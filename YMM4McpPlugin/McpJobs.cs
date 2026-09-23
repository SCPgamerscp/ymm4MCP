using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        private static readonly Regex JobIdPattern = new("^job_[A-Za-z0-9]{8,64}$", RegexOptions.Compiled);
        private static readonly JsonSerializerOptions JobJson = new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };
        private readonly object _jobsLock = new();
        private readonly Dictionary<string, McpJob> _jobs = new();
        private readonly Dictionary<string, string> _jobKeys = new();

        private static string JobsFilePath => Path.Combine(McpSettings.DirectoryPath, "jobs.json");

        private static bool IsNonBlockingPost(string path)
            => path.Equals("/api/project/export", StringComparison.OrdinalIgnoreCase)
               || path.StartsWith("/api/jobs", StringComparison.OrdinalIgnoreCase);

        private async Task<object?> TryRouteJobs(HttpListenerRequest req, string path)
        {
            if (path.Equals("/api/project/export", StringComparison.OrdinalIgnoreCase) && req.HttpMethod == "POST")
                return await EnqueueExport(req);
            if (!path.StartsWith("/api/jobs", StringComparison.OrdinalIgnoreCase))
                return null;
            if (path.Equals("/api/jobs", StringComparison.OrdinalIgnoreCase) || path.Equals("/api/jobs/", StringComparison.OrdinalIgnoreCase))
            {
                if (req.HttpMethod == "GET") return ListJobs();
                if (req.HttpMethod == "POST") return await EnqueueExport(req);
                throw new ArgumentException("Unsupported jobs method");
            }
            var rest = path.Length > 10 ? path[10..].Trim('/') : "";
            var parts = rest.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || !JobIdPattern.IsMatch(parts[0]))
                throw new ArgumentException("job id is invalid");
            string id = parts[0];
            if (parts.Length == 1 && req.HttpMethod == "GET")
                return GetJob(id);
            if (parts.Length == 2 && parts[1].Equals("cancel", StringComparison.OrdinalIgnoreCase) && req.HttpMethod == "POST")
                return CancelJob(id);
            if (parts.Length == 2 && parts[1].Equals("resume", StringComparison.OrdinalIgnoreCase) && req.HttpMethod == "POST")
                return ResumeJob(id);
            return Failure("NOT_FOUND", "Unknown jobs path");
        }

        private object ListJobs()
        {
            lock (_jobsLock)
            {
                return new
                {
                    success = true,
                    jobs = _jobs.Values
                        .OrderByDescending(j => j.UpdatedAt)
                        .Take(50)
                        .Select(DescribeJob)
                        .ToArray()
                };
            }
        }

        private object GetJob(string id)
        {
            lock (_jobsLock)
            {
                if (!_jobs.TryGetValue(id, out var job))
                    return Failure("JOB_NOT_FOUND", "ジョブが見つかりません: " + id);
                return DescribeJob(job);
            }
        }

        private object CancelJob(string id)
        {
            McpJob? job;
            lock (_jobsLock)
            {
                if (!_jobs.TryGetValue(id, out job))
                    return Failure("JOB_NOT_FOUND", "ジョブが見つかりません: " + id);
                if (job.Status is "completed" or "cancelled")
                    return DescribeJob(job);
                job.CancelRequested = true;
            }
            try { job.Cts.Cancel(); } catch { }
            UpdateJob(job, job.Status == "queued" ? "cancelled" : job.Status, message: "キャンセル要求を受け付けました");
            return DescribeJob(job);
        }

        private object ResumeJob(string id)
        {
            McpJob job;
            lock (_jobsLock)
            {
                if (!_jobs.TryGetValue(id, out job!))
                    return Failure("JOB_NOT_FOUND", "ジョブが見つかりません: " + id);
                if (job.Status is "queued" or "running")
                    return DescribeJob(job);
                if (job.Kind != "export")
                    return Failure("JOB_NOT_RESUMABLE", "このジョブ種別は再開できません");
                if (job.Status is not ("failed" or "interrupted" or "cancelled"))
                    return Failure("JOB_NOT_RESUMABLE", "完了済みジョブは再開できません");
            }
            return RestartExport(job);
        }

        private void RestorePersistedJobs()
        {
            try
            {
                if (!File.Exists(JobsFilePath)) return;
                var snapshots = JsonSerializer.Deserialize<List<JobSnapshot>>(File.ReadAllText(JobsFilePath), JobJson);
                if (snapshots == null) return;
                lock (_jobsLock)
                {
                    foreach (var snap in snapshots.TakeLast(30))
                    {
                        if (string.IsNullOrEmpty(snap.Id) || _jobs.ContainsKey(snap.Id)) continue;
                        var job = snap.ToJob();
                        if (job.Status is "queued" or "running")
                        {
                            job.Status = "interrupted";
                            job.Phase = "interrupted";
                            job.Message = "YMM4再起動により中断されました。resumeで再投入できます";
                            job.ErrorCode = "JOB_INTERRUPTED";
                        }
                        _jobs[job.Id] = job;
                        if (!string.IsNullOrEmpty(job.IdempotencyKey))
                            _jobKeys[job.IdempotencyKey] = job.Id;
                    }
                }
            }
            catch (Exception ex) { Log("ジョブ復元に失敗: " + ex.Message); }
        }

        private void CancelAndPersistJobs()
        {
            List<McpJob> running;
            lock (_jobsLock)
            {
                running = _jobs.Values.Where(j => j.Status is "queued" or "running").ToList();
            }
            foreach (var job in running)
            {
                try { job.Cts.Cancel(); } catch { }
                UpdateJob(job, "interrupted", phase: "interrupted", message: "サーバー停止により中断", errorCode: "JOB_INTERRUPTED");
            }
            PersistJobs();
        }

        private void PersistJobs()
        {
            try
            {
                McpSettings.PrepareDirectory();
                JobSnapshot[] snapshots;
                lock (_jobsLock)
                {
                    snapshots = _jobs.Values.OrderByDescending(j => j.UpdatedAt).Take(30).Select(JobSnapshot.From).ToArray();
                }
                string temp = JobsFilePath + ".tmp";
                File.WriteAllText(temp, JsonSerializer.Serialize(snapshots, JobJson));
                File.Move(temp, JobsFilePath, true);
            }
            catch (Exception ex) { Log("ジョブ保存に失敗: " + ex.Message); }
        }

        private McpJob? FindIdempotentJob(string? key)
        {
            if (string.IsNullOrEmpty(key)) return null;
            lock (_jobsLock)
            {
                if (_jobKeys.TryGetValue(key, out var id) && _jobs.TryGetValue(id, out var job)
                    && job.Status is "queued" or "running" or "completed")
                    return job;
            }
            return null;
        }

        private McpJob CreateJob(string kind, Dictionary<string, object?> request, string? idempotencyKey, out bool created)
        {
            lock (_jobsLock)
            {
                if (!string.IsNullOrEmpty(idempotencyKey) && _jobKeys.TryGetValue(idempotencyKey, out var existingId)
                    && _jobs.TryGetValue(existingId, out var existing)
                    && existing.Status is "queued" or "running" or "completed")
                {
                    created = false;
                    return existing;
                }
                var job = new McpJob
                {
                    Id = "job_" + Guid.NewGuid().ToString("N"),
                    Kind = kind,
                    Status = "queued",
                    Phase = "queued",
                    CreatedAt = DateTime.UtcNow.ToString("o"),
                    UpdatedAt = DateTime.UtcNow.ToString("o"),
                    IdempotencyKey = idempotencyKey,
                    Request = request
                };
                _jobs[job.Id] = job;
                if (!string.IsNullOrEmpty(idempotencyKey))
                    _jobKeys[idempotencyKey] = job.Id;
                PruneJobsLocked();
                created = true;
                return job;
            }
        }

        private void PruneJobsLocked()
        {
            if (_jobs.Count <= 80) return;
            foreach (var old in _jobs.Values.Where(j => j.Status is "completed" or "failed" or "cancelled" or "interrupted")
                         .OrderBy(j => j.UpdatedAt).Take(_jobs.Count - 50).ToArray())
            {
                _jobs.Remove(old.Id);
                if (!string.IsNullOrEmpty(old.IdempotencyKey))
                    _jobKeys.Remove(old.IdempotencyKey);
                old.Cts.Dispose();
            }
        }

        private void UpdateJob(McpJob job, string? status = null, string? phase = null, int? progress = null,
            string? message = null, string? error = null, string? errorCode = null, object? result = null, object? checkpoint = null)
        {
            lock (_jobsLock)
            {
                if (status != null) job.Status = status;
                if (phase != null) job.Phase = phase;
                if (progress != null) job.Progress = Math.Clamp(progress.Value, 0, 100);
                if (message != null) job.Message = message;
                if (error != null) job.Error = error;
                if (errorCode != null) job.ErrorCode = errorCode;
                if (result != null) job.Result = result;
                if (checkpoint != null) job.Checkpoint = checkpoint;
                job.UpdatedAt = DateTime.UtcNow.ToString("o");
            }
            PersistJobs();
        }

        private void RunJob(McpJob job, Func<McpJob, CancellationToken, Task> work)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    UpdateJob(job, "running", phase: "waiting_for_edit_lock", progress: 1, message: "編集ロック待ち");
                    await _editGate.WaitAsync(job.Cts.Token);
                    try
                    {
                        if (job.CancelRequested) throw new OperationCanceledException();
                        await work(job, job.Cts.Token);
                    }
                    finally
                    {
                        _editGate.Release();
                    }
                }
                catch (OperationCanceledException)
                {
                    UpdateJob(job, "cancelled", phase: "cancelled", message: "キャンセルされました");
                }
                catch (Exception ex)
                {
                    UpdateJob(job, "failed", phase: "failed",
                        message: ex.InnerException?.Message ?? ex.Message,
                        error: ex.InnerException?.Message ?? ex.Message,
                        errorCode: ex is ArgumentException ? "INVALID_ARGUMENT" : "JOB_FAILED");
                }
            });
        }

        private static object DescribeJob(McpJob job) => new
        {
            success = job.Status is "completed" or "queued" or "running",
            job_id = job.Id,
            kind = job.Kind,
            status = job.Status,
            phase = job.Phase,
            progress = job.Progress,
            message = job.Message,
            error = job.Error,
            error_code = job.ErrorCode,
            created_at = job.CreatedAt,
            updated_at = job.UpdatedAt,
            idempotency_key = job.IdempotencyKey,
            request = job.Request,
            checkpoint = job.Checkpoint,
            result = job.Result,
            cancel_requested = job.CancelRequested
        };

        internal sealed class McpJob
        {
            public string Id { get; init; } = "";
            public string Kind { get; init; } = "";
            public string Status { get; set; } = "queued";
            public string Phase { get; set; } = "queued";
            public int Progress { get; set; }
            public string Message { get; set; } = "";
            public string? Error { get; set; }
            public string? ErrorCode { get; set; }
            public string CreatedAt { get; init; } = "";
            public string UpdatedAt { get; set; } = "";
            public string? IdempotencyKey { get; init; }
            public Dictionary<string, object?> Request { get; init; } = new();
            public object? Result { get; set; }
            public object? Checkpoint { get; set; }
            public bool CancelRequested { get; set; }
            public CancellationTokenSource Cts { get; } = new();
        }

        private sealed class JobSnapshot
        {
            public string Id { get; set; } = "";
            public string Kind { get; set; } = "";
            public string Status { get; set; } = "";
            public string Phase { get; set; } = "";
            public int Progress { get; set; }
            public string Message { get; set; } = "";
            public string? Error { get; set; }
            public string? ErrorCode { get; set; }
            public string CreatedAt { get; set; } = "";
            public string UpdatedAt { get; set; } = "";
            public string? IdempotencyKey { get; set; }
            public Dictionary<string, JsonElement>? Request { get; set; }
            public JsonElement? Result { get; set; }
            public JsonElement? Checkpoint { get; set; }

            public static JobSnapshot From(McpJob job) => new()
            {
                Id = job.Id,
                Kind = job.Kind,
                Status = job.Status,
                Phase = job.Phase,
                Progress = job.Progress,
                Message = job.Message,
                Error = job.Error,
                ErrorCode = job.ErrorCode,
                CreatedAt = job.CreatedAt,
                UpdatedAt = job.UpdatedAt,
                IdempotencyKey = job.IdempotencyKey,
                Request = JsonSerializer.SerializeToElement(job.Request).Deserialize<Dictionary<string, JsonElement>>(),
                Result = job.Result == null ? null : JsonSerializer.SerializeToElement(job.Result),
                Checkpoint = job.Checkpoint == null ? null : JsonSerializer.SerializeToElement(job.Checkpoint)
            };

            public McpJob ToJob()
            {
                var request = new Dictionary<string, object?>();
                if (Request != null)
                {
                    foreach (var kv in Request)
                        request[kv.Key] = kv.Value.ValueKind == JsonValueKind.Undefined ? null : kv.Value;
                }
                return new McpJob
                {
                    Id = Id,
                    Kind = Kind,
                    Status = Status,
                    Phase = Phase,
                    Progress = Progress,
                    Message = Message,
                    Error = Error,
                    ErrorCode = ErrorCode,
                    CreatedAt = CreatedAt,
                    UpdatedAt = UpdatedAt,
                    IdempotencyKey = IdempotencyKey,
                    Request = request,
                    Result = Result is JsonElement r && r.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null ? r : null,
                    Checkpoint = Checkpoint is JsonElement c && c.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null ? c : null
                };
            }
        }
    }
}
