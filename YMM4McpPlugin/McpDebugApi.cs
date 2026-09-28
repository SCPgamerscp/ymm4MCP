using System;
using System.Globalization;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        // ── デバッグ ──────────────────────────────────────────

        private object DebugViewModel()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var props = vm.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name, isCommand = typeof(ICommand).IsAssignableFrom(p.PropertyType) })
                    .OrderBy(p => p.name).ToArray();
                return (object)new { vmTypeName = vm.GetType().FullName, properties = props };
            });
        }

        private object DebugTimeline()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var props = tvm.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name }).OrderBy(p => p.name).ToArray();
                return (object)new { timelineVmType = tvm.GetType().Name, properties = props };
            });
        }

        private object DebugItems()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { error = "Items取得失敗" };
                var result = new List<object>();
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    var props = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => { object? v = null; try { v = p.GetValue(item)?.ToString(); } catch { } return new { name = p.Name, value = v }; })
                        .Where(p => p.value != null).ToArray();
                    result.Add(new { typeName = item.GetType().Name, properties = props });
                    break;
                }
                return (object)result;
            });
        }

        private object DebugScene()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var docVMs = GetPropEnum(vm, "DocumentViewModels");
                if (docVMs == null) return (object)new { error = "DocumentViewModels取得失敗" };
                var result = new List<object>();
                foreach (var docVM in docVMs)
                {
                    var sceneVM = GetPropObj(docVM, "ViewModel");
                    result.Add(new
                    {
                        docVMType = docVM.GetType().Name,
                        sceneVMType = sceneVM?.GetType().Name,
                        sceneVMProps = sceneVM?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    });
                    break;
                }
                return (object)result;
            });
        }

        private object DebugVoiceTypes()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                // ロード済みアセンブリからVoiceItem系のクラスを探す
                var voiceTypes = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => { try { return a.GetTypes(); } catch { return Array.Empty<Type>(); } })
                    .Where(t => t.Name.Contains("Voice") && t.Name.Contains("Item") && !t.IsInterface && !t.IsAbstract)
                    .Select(t => new { fullName = t.FullName, ctors = t.GetConstructors().Select(c => string.Join(", ", c.GetParameters().Select(p => $"{p.ParameterType.Name} {p.Name}"))).ToArray() })
                    .ToArray();

                return (object)new { voiceTypes };
            });
        }

        private object DebugTimelineMethods()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                // timelineフィールドのメソッドを確認
                var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
                var timelineObj = timelineField?.GetValue(tvm);
                var sceneField = tvm.GetType().GetField("scene", BindingFlags.NonPublic | BindingFlags.Instance);
                var sceneObj = sceneField?.GetValue(tvm);

                var timelineMethods = timelineObj?.GetType()
                    .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName)
                    .Select(m => $"{m.Name}({string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name))})")
                    .OrderBy(s => s).ToArray();

                var sceneMethods = sceneObj?.GetType()
                    .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName)
                    .Select(m => $"{m.Name}({string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name))})")
                    .OrderBy(s => s).ToArray();

                return (object)new { timelineMethods, sceneMethods };
            });
        }

        private object DebugMenuItem()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                var menuVM = GetPropObj(tvm, "AddVoiceItemTemplateContextMenuViewModel");
                var menuItems = menuVM?.GetType().GetProperty("Items")?.GetValue(menuVM) as System.Collections.IEnumerable;

                // 再帰的にコマンドを探す
                var result = new List<object>();
                void Explore(object item, int depth)
                {
                    if (depth > 4) return;
                    var t = item.GetType();
                    var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new { p.Name, TypeName = p.PropertyType.Name, IsCommand = typeof(ICommand).IsAssignableFrom(p.PropertyType) })
                        .ToArray();
                    result.Add(new { depth, itemType = t.Name, props });

                    // 子Items再帰
                    var childItems = t.GetProperty("Items")?.GetValue(item) as System.Collections.IEnumerable;
                    if (childItems != null)
                        foreach (var child in childItems) { Explore(child, depth + 1); break; } // 1件のみ
                }
                if (menuItems != null)
                    foreach (var item in menuItems) { Explore(item, 0); break; }

                return (object)result;
            });
        }

        private object DebugVoiceCmd()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                // AddVoiceItemCommandParameterの中身
                var paramProp = tvm.GetType().GetProperty("AddVoiceItemCommandParameter");
                var paramVal = paramProp?.GetValue(tvm);
                var innerVal = paramVal?.GetType().GetProperty("Value")?.GetValue(paramVal);

                // AddVoiceItemTemplateContextMenuViewModel.Items の中身
                var menuVM = GetPropObj(tvm, "AddVoiceItemTemplateContextMenuViewModel");
                object[]? menuItems = null;
                if (menuVM != null)
                {
                    var items = menuVM.GetType().GetProperty("Items")?.GetValue(menuVM) as System.Collections.IEnumerable;
                    menuItems = items?.Cast<object>().Select(item => new
                    {
                        type = item.GetType().Name,
                        props = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray()
                    } as object).ToArray();
                }

                // CurrentCharacterの中身
                var charProp = tvm.GetType().GetProperty("CurrentCharacter");
                var charVal = charProp?.GetValue(tvm);
                var charInner = charVal?.GetType().GetProperty("Value")?.GetValue(charVal);

                // Characters一覧
                var charsEnum = GetPropEnum(tvm, "Characters");
                var charNames = charsEnum?.Cast<object>().Select(c =>
                {
                    var t = c.GetType();
                    return t.GetProperties().Where(p => p.CanRead).Select(p =>
                    {
                        string? v = null; try { v = p.GetValue(c)?.ToString(); } catch { }
                        return new { p.Name, val = v };
                    }).ToArray();
                }).ToArray();

                return (object)new
                {
                    paramType = paramVal?.GetType().Name,
                    innerValType = innerVal?.GetType().Name,
                    innerValProps = innerVal?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    menuItems,
                    charInnerType = charInner?.GetType().Name,
                    charNames
                };
            });
        }

        private object DebugItemToolBar()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                var toolbar = GetPropObj(tvm!, "ToolBar");
                var itemToolBar = GetPropObj(toolbar!, "ItemToolBar");
                if (itemToolBar == null) return (object)new { error = "ItemToolBar取得失敗" };

                var props = itemToolBar.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name }).OrderBy(p => p.name).ToArray();
                return (object)new { type = itemToolBar.GetType().FullName, props };
            });
        }

        private object DebugToolBar()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var toolbar = GetPropObj(tvm, "ToolBar");
                if (toolbar == null) return (object)new { error = "ToolBar取得失敗" };

                var props = toolbar.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name }).OrderBy(p => p.name).ToArray();
                var methods = toolbar.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName).Select(m => m.Name).OrderBy(n => n).ToArray();
                return (object)new { toolbarType = toolbar.GetType().FullName, props, methods };
            });
        }

        private object DebugSceneFields()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                var sceneField = tvm.GetType().GetField("scene", BindingFlags.NonPublic | BindingFlags.Instance);
                var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
                var sceneObj = sceneField?.GetValue(tvm);
                var timelineObj = timelineField?.GetValue(tvm);

                return (object)new
                {
                    sceneType = sceneObj?.GetType().FullName,
                    sceneProps = sceneObj?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    timelineType = timelineObj?.GetType().FullName,
                    timelineProps = timelineObj?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                };
            });
        }

    }
}
