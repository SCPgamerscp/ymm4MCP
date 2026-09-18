using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Loader;
using System.Windows;
using System.Windows.Input;

namespace YMM4McpPlugin
{
    public class McpViewModel : INotifyPropertyChanged
    {
        // A tool view may be opened more than once. All views observe one listener.
        internal static readonly McpHttpServer Server = new();
        private static string _startupError = "";
        private bool _attached;
        private string _log = "";
        private int _port = Server.Settings.Port;
        private bool _autoStart = Server.Settings.AutoStart;
        private bool _allowAdvanced = Server.Settings.AllowAdvanced;

        static McpViewModel()
        {
            if (Application.Current != null)
                Application.Current.Exit += (_, _) => Server.Stop();
            var context = AssemblyLoadContext.GetLoadContext(typeof(McpViewModel).Assembly);
            if (context != null) context.Unloading += _ => Server.Stop();
            if (Server.Settings.AutoStart)
                try { Server.Start(); } catch (Exception ex) { _startupError = ex.Message; }
        }

        internal static void InitializePlugin() { _ = Server; }
        public string Log { get => _log; private set { _log = value; OnPropertyChanged(); } }
        public bool IsRunning => Server.IsRunning;
        public bool CanConfigure => !IsRunning;
        public string StatusText => IsRunning ? "起動中" : "停止";
        public string BaseUrl => Server.BaseUrl;
        public string ConnectionPath => McpSettings.ConnectionPath;
        public int Port { get => _port; set { _port = value; OnPropertyChanged(); } }
        public bool AutoStart { get => _autoStart; set { _autoStart = value; OnPropertyChanged(); } }
        public bool AllowAdvanced { get => _allowAdvanced; set { _allowAdvanced = value; OnPropertyChanged(); } }

        public ICommand StartCommand => new RelayCommand(_ => RunSafely(() =>
        {
            SaveSettings();
            Server.Start();
        }), _ => !IsRunning);
        public ICommand StopCommand => new RelayCommand(_ => RunSafely(Server.Stop), _ => IsRunning);
        public ICommand SaveSettingsCommand => new RelayCommand(_ => RunSafely(SaveSettings), _ => !IsRunning);
        public ICommand ClearLogCommand => new RelayCommand(_ => Log = "");

        private void SaveSettings()
        {
            if (IsRunning) throw new InvalidOperationException("設定変更前にサーバーを停止してください");
            var candidate = new McpSettings { Port = Port, AutoStart = AutoStart, AllowAdvanced = AllowAdvanced };
            candidate.Save();
            Server.Settings.Port = candidate.Port;
            Server.Settings.AutoStart = candidate.AutoStart;
            Server.Settings.AllowAdvanced = candidate.AllowAdvanced;
            Refresh();
        }

        private void RunSafely(Action action)
        {
            try { action(); } catch (Exception ex) { AddLog(ex.Message); }
            finally { Refresh(); }
        }

        public void Attach()
        {
            if (_attached) return;
            _attached = true;
            Server.LogMessage += AddLog;
            Server.StateChanged += Refresh;
            Port = Server.Settings.Port;
            AutoStart = Server.Settings.AutoStart;
            AllowAdvanced = Server.Settings.AllowAdvanced;
            if (_startupError.Length > 0) AddLog(_startupError);
            Refresh();
        }

        public void Detach()
        {
            if (!_attached) return;
            _attached = false;
            Server.LogMessage -= AddLog;
            Server.StateChanged -= Refresh;
            // Closing the tool window does not interrupt an in-flight MCP task.
        }

        private void Refresh() => OnUI(() =>
        {
            OnPropertyChanged(nameof(IsRunning));
            OnPropertyChanged(nameof(CanConfigure));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(BaseUrl));
            CommandManager.InvalidateRequerySuggested();
        });

        private void AddLog(string message) => OnUI(() =>
        {
            Log += message + "\n";
            if (Log.Length > 5000) Log = Log[^4000..];
        });

        private static void OnUI(Action action)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.HasShutdownStarted) return;
            if (dispatcher.CheckAccess()) action();
            else dispatcher.InvokeAsync(action);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Func<object?, bool>? _canExecute;
        public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
        { _execute = execute; _canExecute = canExecute; }
        public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object? parameter) => _execute(parameter);
        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
