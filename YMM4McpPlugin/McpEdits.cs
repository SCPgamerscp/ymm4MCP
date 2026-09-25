using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;
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
        private static readonly Regex CheckpointIdPattern = new("^cp_[A-Za-z0-9]{8,64}$", RegexOptions.Compiled);
        private readonly object _editsLock = new();
        private List<EditBindingRecord> _editBindings = new();
        private List<CheckpointRecord> _checkpoints = new();

        private static string EditsFilePath => Path.Combine(McpSettings.DirectoryPath, "edits.json");
        private static string CheckpointsFilePath => Path.Combine(McpSettings.DirectoryPath, "checkpoints.json");
        private static string CheckpointBackupDir => Path.Combine(McpSettings.DirectoryPath, "checkpoints");

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
            if (rest.Equals("checkpoint", StringComparison.OrdinalIgnoreCase))
            {
                if (req.HttpMethod != "POST") throw new ArgumentException("Unsupported checkpoint method");
                return await CreateCheckpoint(req);
            }
            if (rest.Equals("checkpoints", StringComparison.OrdinalIgnoreCase)
                || rest.Equals("checkpoints/", StringComparison.OrdinalIgnoreCase))
            {
                if (req.HttpMethod != "GET") throw new ArgumentException("Unsupported checkpoints method");
                return ListCheckpoints();
            }
            if (rest.StartsWith("checkpoints/", StringComparison.OrdinalIgnoreCase))
            {
                string id = rest[12..].Trim('/');
                if (!CheckpointIdPattern.IsMatch(id)) throw new ArgumentException("checkpoint id is invalid");
                if (req.HttpMethod != "GET") throw new ArgumentException("Unsupported checkpoint method");
                return GetCheckpoint(id);
            }
            if (rest.Equals("rollback", StringComparison.OrdinalIgnoreCase))
            {
                if (req.HttpMethod != "POST") throw new ArgumentException("Unsupported rollback method");
                return await RollbackCheckpoint(req);
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
                checkpoints = SnapshotCheckpoints().Select(DescribeCheckpoint).ToArray(),
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

        private List<string> SnapshotItemIds()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var ids = new List<string>();
                var vm = GetMainViewModel();
                if (vm == null) return ids;
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return ids;
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return ids;
                foreach (var iv in rawItems)
                {
                    var info = ReadItemInfo(iv);
                    string id = GetItemIdentity(info.item).id;
                    if (!string.IsNullOrEmpty(id)) ids.Add(id);
                }
                return ids;
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
                var existing = _editBindings.FirstOrDefault(b => b.IdempotencyKey == key);
                if (existing != null && !string.Equals(existing.PlanHash, hash, StringComparison.Ordinal))
                    return Failure("IDEMPOTENCY_KEY_CONFLICT", "idempotency_key is already bound to a different EditPlan");
                var updated = _editBindings.Where(b => b.IdempotencyKey != key).Select(b => b.Clone()).ToList();
                updated.Add(record);
                if (updated.Count > 40)
                    updated = updated.TakeLast(30).ToList();
                try
                {
                    PersistEditBindings(updated);
                    _editBindings = updated;
                }
                catch (Exception ex)
                {
                    Log("EditPlanバインディング保存に失敗: " + ex.Message);
                    return Failure("BINDINGS_NOT_SAVED", "EditPlanバインディングを保存できませんでした");
                }
            }
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

        private static void PersistEditBindings(List<EditBindingRecord> records)
        {
            McpSettings.PrepareDirectory();
            string temp = EditsFilePath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(records, EditJson));
            File.Move(temp, EditsFilePath, true);
        }

        private void ForgetEditBindings(IEnumerable<string> itemIds)
        {
            var removed = new HashSet<string>(itemIds.Where(id => !string.IsNullOrEmpty(id)));
            if (removed.Count == 0) return;
            lock (_editsLock)
            {
                var updated = _editBindings.Select(b => b.Clone()).ToList();
                foreach (var record in updated)
                    record.Items.RemoveAll(i => removed.Contains(i.ItemId));
                try
                {
                    PersistEditBindings(updated);
                    _editBindings = updated;
                }
                catch (Exception ex)
                {
                    Log("EditPlanバインディング更新に失敗: " + ex.Message);
                }
            }
        }

        private static object DescribeBinding(EditBindingRecord record) => new
        {
            idempotency_key = record.IdempotencyKey,
            plan_hash = record.PlanHash,
            updated_at = record.UpdatedAt,
            items = record.Items.Select(i => new { id = i.Id, item_id = i.ItemId, revision = i.Revision }).ToArray()
        };

        private async Task<object> CreateCheckpoint(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            string reason = GetStr(body, "reason", "");
            if (reason.Length > 200)
                throw new ArgumentException("reason must be at most 200 characters");
            bool backup = GetBool(body, "backup", false);
            string id = "cp_" + Guid.NewGuid().ToString("N");
            var itemIds = SnapshotItemIds();
            string? projectPath = CurrentProjectPath();
            string? backupPath = null;
            string? backupWarning = null;
            if (backup)
            {
                if (string.IsNullOrEmpty(projectPath) || !File.Exists(projectPath))
                    backupWarning = "NO_PROJECT_FILE";
                else
                {
                    try
                    {
                        Directory.CreateDirectory(CheckpointBackupDir);
                        backupPath = Path.Combine(CheckpointBackupDir, id + ".ymmp");
                        File.Copy(projectPath, backupPath);
                    }
                    catch (Exception ex)
                    {
                        backupPath = null;
                        backupWarning = "BACKUP_FAILED:" + ex.Message;
                    }
                }
            }
            var record = new CheckpointRecord
            {
                Id = id,
                CreatedAt = DateTime.UtcNow.ToString("o"),
                Reason = reason,
                ProjectPath = projectPath,
                BackupPath = backupPath,
                ItemIds = itemIds
            };
            List<CheckpointRecord> kept;
            List<CheckpointRecord> pruned;
            lock (_editsLock)
            {
                kept = _checkpoints.Select(c => c.Clone()).ToList();
                kept.Add(record);
                pruned = kept.Count <= 20 ? new List<CheckpointRecord>() : kept.Take(kept.Count - 20).ToList();
                if (pruned.Count > 0)
                    kept = kept.TakeLast(20).ToList();
            }
            try
            {
                PersistCheckpoints(kept);
            }
            catch (Exception ex)
            {
                Log("チェックポイント保存に失敗: " + ex.Message);
                TryDeleteBackup(backupPath);
                return Failure("CHECKPOINT_NOT_SAVED", "チェックポイントを保存できませんでした");
            }
            lock (_editsLock)
                _checkpoints = kept;
            foreach (var old in pruned)
                TryDeleteBackup(old.BackupPath);
            return new
            {
                success = true,
                checkpoint_id = record.Id,
                created_at = record.CreatedAt,
                reason = record.Reason,
                item_count = record.ItemIds.Count,
                project_path = record.ProjectPath,
                backup_path = record.BackupPath,
                backup_warning = backupWarning,
                note = "Rollback deletes items added after this snapshot. Deleted items and property edits are not restored unless you open backup_path."
            };
        }

        private object ListCheckpoints()
        {
            lock (_editsLock)
            {
                return new
                {
                    success = true,
                    checkpoints = _checkpoints.Select(DescribeCheckpoint).Reverse().ToArray()
                };
            }
        }

        private object GetCheckpoint(string id)
        {
            lock (_editsLock)
            {
                var record = _checkpoints.FirstOrDefault(c => c.Id == id);
                if (record == null) return Failure("CHECKPOINT_NOT_FOUND", "チェックポイントが見つかりません: " + id);
                return DescribeCheckpoint(record, includeItems: true);
            }
        }

        private async Task<object> RollbackCheckpoint(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            string id = GetStr(body, "checkpoint_id", "");
            if (!CheckpointIdPattern.IsMatch(id))
                throw new ArgumentException("checkpoint_id is invalid");
            CheckpointRecord? record;
            lock (_editsLock)
                record = _checkpoints.FirstOrDefault(c => c.Id == id);
            if (record == null)
                return Failure("CHECKPOINT_NOT_FOUND", "チェックポイントが見つかりません: " + id);

            var kept = new HashSet<string>(record.ItemIds);
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return Failure("NO_MAIN_VIEW_MODEL", "VM失敗");
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return Failure("NO_TIMELINE", "TVM失敗");
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return Failure("ITEMS_UNAVAILABLE", "Items失敗");
                var currentIds = new List<string>();
                var toRemove = new List<object>();
                foreach (var iv in rawItems)
                {
                    var info = ReadItemInfo(iv);
                    string itemId = GetItemIdentity(info.item).id;
                    currentIds.Add(itemId);
                    if (!kept.Contains(itemId))
                        toRemove.Add(info.item);
                }
                var missingSnapshot = record.ItemIds.Where(itemId => !currentIds.Contains(itemId)).ToArray();
                if (toRemove.Count == 0)
                {
                    return new
                    {
                        success = true,
                        checkpoint_id = record.Id,
                        removed = 0,
                        item_ids = Array.Empty<string>(),
                        missing_from_snapshot = missingSnapshot,
                        backup_path = record.BackupPath,
                        note = missingSnapshot.Length == 0
                            ? "No items were added after the checkpoint."
                            : "Some snapshot items are gone and cannot be recreated. Open backup_path for a full restore if one exists."
                    };
                }
                var deleted = TryRemoveTimelineItems(tvm, toRemove);
                if (!deleted.ok) return deleted.payload;
                return new
                {
                    success = true,
                    checkpoint_id = record.Id,
                    removed = toRemove.Count,
                    item_ids = toRemove.Select(item => GetItemIdentity(item).id).ToArray(),
                    missing_from_snapshot = missingSnapshot,
                    backup_path = record.BackupPath,
                    note = missingSnapshot.Length == 0
                        ? "Items added after the checkpoint were deleted. Property mutations on surviving items were not reverted."
                        : "Some snapshot items are gone and cannot be recreated. Open backup_path for a full restore if one exists."
                };
            });
        }

        private string? CurrentProjectPath()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return null;
                var model = GetMainModel(vm);
                string? path = GetPropValue(vm, "ProjectFilePath")?.ToString()
                    ?? (model == null ? null : GetPropObj(model, "ProjectFilePath")?.ToString());
                return string.IsNullOrWhiteSpace(path) ? null : path;
            });
        }

        private List<CheckpointRecord> SnapshotCheckpoints()
        {
            lock (_editsLock)
                return _checkpoints.Select(c => c.Clone()).ToList();
        }

        private void RestoreCheckpoints()
        {
            try
            {
                if (!File.Exists(CheckpointsFilePath)) return;
                var records = JsonSerializer.Deserialize<List<CheckpointRecord>>(File.ReadAllText(CheckpointsFilePath), EditJson);
                if (records == null) return;
                lock (_editsLock)
                    _checkpoints = records.TakeLast(20).ToList();
            }
            catch (Exception ex) { Log("チェックポイント復元に失敗: " + ex.Message); }
        }

        private static void PersistCheckpoints(List<CheckpointRecord> records)
        {
            McpSettings.PrepareDirectory();
            string temp = CheckpointsFilePath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(records, EditJson));
            File.Move(temp, CheckpointsFilePath, true);
        }

        private static void TryDeleteBackup(string? path)
        {
            if (string.IsNullOrEmpty(path)) return;
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        private static object DescribeCheckpoint(CheckpointRecord record) => DescribeCheckpoint(record, false);

        private static object DescribeCheckpoint(CheckpointRecord record, bool includeItems) => new
        {
            checkpoint_id = record.Id,
            created_at = record.CreatedAt,
            reason = record.Reason,
            item_count = record.ItemIds.Count,
            project_path = record.ProjectPath,
            backup_path = record.BackupPath,
            item_ids = includeItems ? record.ItemIds.ToArray() : null
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

        private sealed class CheckpointRecord
        {
            public string Id { get; set; } = "";
            public string CreatedAt { get; set; } = "";
            public string Reason { get; set; } = "";
            public string? ProjectPath { get; set; }
            public string? BackupPath { get; set; }
            public List<string> ItemIds { get; set; } = new();

            public CheckpointRecord Clone() => new()
            {
                Id = Id,
                CreatedAt = CreatedAt,
                Reason = Reason,
                ProjectPath = ProjectPath,
                BackupPath = BackupPath,
                ItemIds = ItemIds.ToList()
            };
        }
    }
}
