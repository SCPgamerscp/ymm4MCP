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
        public bool AutoStart { get; set; } = false;
        public bool AllowAdvanced { get; set; } = false;
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
            if (!File.Exists(SettingsPath)) return new McpSettings();
            var settings = JsonSerializer.Deserialize<McpSettings>(File.ReadAllText(SettingsPath))
                ?? throw new InvalidDataException("MCP設定ファイルが空です");
            settings.Validate();
            return settings;
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
            string temp = SettingsPath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(this));
            File.Move(temp, SettingsPath, true);
        }
    }
}
