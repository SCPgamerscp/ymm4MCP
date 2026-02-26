using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace YMM4McpPlugin
{
    public class McpViewModel : INotifyPropertyChanged
    {
        private readonly McpHttpServer _server = new();
        private string _log = "サーバー停止中\n";
        private string _statusText = "● 停止";
        private bool _isRunning = false;

        public string Log
        {
            get => _log;
            set { _log = value; OnPropertyChanged(); }
        }

        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(); }
        }

        public bool IsRunning
        {
            get => _isRunning;
            set { _isRunning = value; OnPropertyChanged(); }
        }

        public string BaseUrl => _server.BaseUrl;

        public ICommand StartCommand => new RelayCommand(_ =>
        {
            _server.Start();
            IsRunning = true;
            StatusText = "● 起動中";
            AddLog($"MCPサーバー起動: {_server.BaseUrl}");
        }, _ => !_isRunning);

        public ICommand StopCommand => new RelayCommand(_ =>
        {
            _server.Stop();
            IsRunning = false;
            StatusText = "● 停止";
            AddLog("MCPサーバー停止");
        }, _ => _isRunning);

        public ICommand ClearLogCommand => new RelayCommand(_ =>
        {
            Log = "";
        });

        public McpViewModel()
        {
            _server.LogMessage += msg => AddLog(msg);
        }

        private void AddLog(string msg)
        {
            // UIスレッドで更新
            System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                Log += msg + "\n";
                // 最大5000文字に制限
                if (Log.Length > 5000)
                    Log = "...(省略)...\n" + Log[^4000..];
            });
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
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object? parameter) => _execute(parameter);
        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
