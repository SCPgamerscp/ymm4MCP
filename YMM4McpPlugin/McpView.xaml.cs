using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace YMM4McpPlugin
{
    public partial class McpView : UserControl
    {
        public McpView()
        {
            InitializeComponent();
            Loaded += (_, _) => (DataContext as McpViewModel)?.Attach();
            Unloaded += (_, _) => (DataContext as McpViewModel)?.Detach();
            DataContextChanged += (_, e) =>
            {
                (e.OldValue as McpViewModel)?.Detach();
                if (IsLoaded) (e.NewValue as McpViewModel)?.Attach();
            };
        }

        private void OpenRepository(object sender, RequestNavigateEventArgs e)
        {
            e.Handled = true;
            try
            {
                Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"リポジトリを開けませんでした: {ex.Message}", "YMM4 MCP");
            }
        }
    }
}
