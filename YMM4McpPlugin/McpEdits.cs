using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        private static readonly JsonSerializerOptions EditJson = new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };
        private readonly object _editsLock = new();
        private List<EditBindingRecord> _editBindings = new();

        private static string EditsFilePath => Path.Combine(McpSettings.DirectoryPath, "edits.json");

        private async Task<object?> TryRouteEdits(HttpListenerRequest req, string path)
        {
            if (!path.StartsWith("/api/edits", StringComparison.OrdinalIgnoreCase))
                return null;
            var rest = path.Equals("/api/edits", StringComparison.OrdinalIgnoreCase)
                       || path.Equals("/api/edits/", StringComparison.OrdinalIgnoreCase)
                ? ""
                : path[10..].Trim('/');
            if (rest.Length == 0 || rest.Equals("state", StringComparison.OrdinalIgnoreCase))
            {
                if (req.HttpMethod != "GET") throw new ArgumentException("Unsupported edits method");
                return GetEditState();
            }
            if (rest.Equals("bindings", StringComparison.OrdinalIgnoreCase))
            {
                if (req.HttpMethod == "GET") return ListEditBindings(req);
                if (req.HttpMethod == "POST") return await SaveEditBindings(req);
                throw new ArgumentException("Unsupported bindings method");
            }
            return Failure("NOT_FOUND", "Unknown edits path");
        }

        private object GetEditState()
        {
            var items = SnapshotTimelineItems();
            return new
            {
                success = true,
                scenes = new[]
                {
                    new { id = "timeline", duration_policy = "fit_content", items }
                },
                item_count = items.Count,
                bindings = SnapshotBindings().Select(DescribeBinding).ToArray(),
                note = "Host timeline has no scene labels; items are returned as one scene."
            };
        }

        private List<object> SnapshotTimelineItems()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return new List<object>();
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return new List<object>();
                var rawItems = GetPropEnum(tvm, "Items");
                var items = new List<object>();
                if (rawItems == null) return items;
                foreach (var iv in rawItems)
                {
                    var info = ReadItemInfo(iv);
                    var identity = GetItemIdentity(info.item);
                    items.Add(new
                    {
                        item_id = identity.id,
                        revision = GetItemRevision(info.item),
                        identity_persistent = identity.persistent,
                        layer = info.layer,
                        frame = info.frame,
                        length = info.length,
                        endFrame = (long)info.frame + info.length,
                        type = info.type,
                        text = info.text,
                    });
                }
                return items;
            });
        }

        private object ListEditBindings(HttpListenerRequest req)
        {
            string key = req.QueryString["idempotency_key"] ?? req.QueryString["key"] ?? "";
            lock (_editsLock)
            {
                var records = string.IsNullOrEmpty(key)
                    ? _editBindings
                    : _editBindings.Where(b => b.IdempotencyKey == key).ToList();
                return new { success = true, bindings = records.Select(DescribeBinding).ToArray() };
            }
        }

        private async Task<object> SaveEditBindings(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            string key = GetStr(body, "idempotency_key", "");
            if (key.Length is < 1 or > 128 || key.Trim() != key)
                throw new ArgumentException("idempotency_key must be a 1..128 character string");
            string hash = GetStr(body, "plan_hash", "");
            if (hash.Length is < 16 or > 128)
                throw new ArgumentException("plan_hash is required");
            if (!body.TryGetValue("items", out var itemsEl) || itemsEl.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("items must be an array");
            var items = new List<EditBoundItem>();
            foreach (var el in itemsEl.EnumerateArray())
            {
                if (el.ValueKind != JsonValueKind.Object) throw new ArgumentException("binding items must be objects");
                string id = el.TryGetProperty("id", out var idEl) && idEl.ValueKind == JsonValueKind.String ? idEl.GetString() ?? "" : "";
                string itemId = el.TryGetProperty("item_id", out var itemEl) && itemEl.ValueKind == JsonValueKind.String ? itemEl.GetString() ?? "" : "";
                if (id.Length == 0 || itemId.Length == 0)
                    throw new ArgumentException("binding items need id and item_id");
                string? revision = el.TryGetProperty("revision", out var revEl) && revEl.ValueKind == JsonValueKind.String
                    ? revEl.GetString() : null;
                items.Add(new EditBoundItem { Id = id, ItemId = itemId, Revision = revision });
            }
            var record = new EditBindingRecord
            {
                IdempotencyKey = key,
                PlanHash = hash,
                UpdatedAt = DateTime.UtcNow.ToString("o"),
                Items = items
            };
            lock (_editsLock)
            {
                _editBindings.RemoveAll(b => b.IdempotencyKey == key);
                _editBindings.Add(record);
                if (_editBindings.Count > 40)
                    _editBindings = _editBindings.TakeLast(30).ToList();
            }
            PersistEditBindings();
            return new { success = true, idempotency_key = key, item_count = items.Count, plan_hash = hash };
        }

        private List<EditBindingRecord> SnapshotBindings()
        {
            lock (_editsLock)
                return _editBindings.Select(b => b.Clone()).ToList();
        }

        private void RestoreEditBindings()
        {
            try
            {
                if (!File.Exists(EditsFilePath)) return;
                var records = JsonSerializer.Deserialize<List<EditBindingRecord>>(File.ReadAllText(EditsFilePath), EditJson);
                if (records == null) return;
                lock (_editsLock)
                    _editBindings = records.TakeLast(30).ToList();
            }
            catch (Exception ex) { Log("EditPlanバインディング復元に失敗: " + ex.Message); }
        }

        private void PersistEditBindings()
        {
            try
            {
                McpSettings.PrepareDirectory();
                List<EditBindingRecord> records;
                lock (_editsLock)
                    records = _editBindings.Select(b => b.Clone()).ToList();
                string temp = EditsFilePath + ".tmp";
                File.WriteAllText(temp, JsonSerializer.Serialize(records, EditJson));
                File.Move(temp, EditsFilePath, true);
            }
            catch (Exception ex) { Log("EditPlanバインディング保存に失敗: " + ex.Message); }
        }

        private static object DescribeBinding(EditBindingRecord record) => new
        {
            idempotency_key = record.IdempotencyKey,
            plan_hash = record.PlanHash,
            updated_at = record.UpdatedAt,
            items = record.Items.Select(i => new { id = i.Id, item_id = i.ItemId, revision = i.Revision }).ToArray()
        };

        private sealed class EditBoundItem
        {
            public string Id { get; set; } = "";
            public string ItemId { get; set; } = "";
            public string? Revision { get; set; }
        }

        private sealed class EditBindingRecord
        {
            public string IdempotencyKey { get; set; } = "";
            public string PlanHash { get; set; } = "";
            public string UpdatedAt { get; set; } = "";
            public List<EditBoundItem> Items { get; set; } = new();

            public EditBindingRecord Clone() => new()
            {
                IdempotencyKey = IdempotencyKey,
                PlanHash = PlanHash,
                UpdatedAt = UpdatedAt,
                Items = Items.Select(i => new EditBoundItem { Id = i.Id, ItemId = i.ItemId, Revision = i.Revision }).ToList()
            };
        }
    }
}
