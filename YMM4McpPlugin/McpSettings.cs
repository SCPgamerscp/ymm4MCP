using System;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.Json;

namespace YMM4McpPlugin
{
    public sealed class McpSettings
    {
        public int Port { get; set; } = 8765;
        public bool AutoStart { get; set; } = true;
        public bool AllowAdvanced { get; set; } = false;
        // A corrupt settings file must not prevent the plugin from loading or
        // silently enable the HTTP listener / advanced reflection API.
        [System.Text.Json.Serialization.JsonIgnore]
        public string LoadWarning { get; set; } = "";
        public static string DirectoryPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "YMM4MCP");
        public static string ConnectionPath => Path.Combine(DirectoryPath, "connection.json");
        private static string SettingsPath => Path.Combine(DirectoryPath, "settings.json");

        internal static void PrepareDirectory()
        {
            var directory = Directory.CreateDirectory(DirectoryPath);
            // The connection file is a bearer credential. Do not inherit broad ACLs.
            var user = WindowsIdentity.GetCurrent().User
                ?? throw new InvalidOperationException("Windows user SID unavailable");
            var security = new DirectorySecurity();
            security.SetAccessRuleProtection(true, false);
            security.SetOwner(user);
            security.AddAccessRule(new FileSystemAccessRule(user, FileSystemRights.FullControl,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                PropagationFlags.None, AccessControlType.Allow));
            directory.SetAccessControl(security);
        }

        public static McpSettings Load()
        {
            try
            {
                // File.Exists can return false for access-denied errors, which
                // would incorrectly auto-start with defaults. Read directly.
                var settings = JsonSerializer.Deserialize<McpSettings>(SettingsFileSafety.ReadValidated(SettingsPath))
                    ?? throw new InvalidDataException("MCP設定ファイルが空です");
                settings.Validate();
                return settings;
            }
            catch (FileNotFoundException) { return new McpSettings(); }
            catch (DirectoryNotFoundException) { return new McpSettings(); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or JsonException or InvalidDataException
                                       or ArgumentOutOfRangeException)
            {
                // Preserve the damaged file for diagnosis. Do not auto-start on
                // fallback settings, and never enable dangerous APIs.
                return new McpSettings
                {
                    AutoStart = false,
                    AllowAdvanced = false,
                    LoadWarning = "MCP設定を読み込めません。自動起動を無効化しました。設定ファイルを確認し、必要ならUIから保存し直してください: " + ex.Message
                };
            }
        }

        public void Validate()
        {
            if (Port < 1024 || Port > 65535)
                throw new ArgumentOutOfRangeException(nameof(Port), "ポートは1024〜65535で指定してください");
        }

        public void Save()
        {
            Validate();
            PrepareDirectory();
            SettingsFileSafety.WriteAtomic(SettingsPath, JsonSerializer.Serialize(this));
        }
    }
}
