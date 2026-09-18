using System.Windows.Controls;

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
    }
}
