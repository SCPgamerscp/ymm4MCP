using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace YMM4McpPlugin
{
    /// <summary>IsRunning の bool -> ブラシ変換 (staticシングルトン)</summary>
    public class BoolToBrushConverter : IValueConverter
    {
        public static readonly BoolToBrushConverter Instance = new();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is true
                ? new SolidColorBrush(Color.FromRgb(0xa6, 0xe3, 0xa1))  // 緑
                : new SolidColorBrush(Color.FromRgb(0xf3, 0x8b, 0xa8)); // 赤

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
