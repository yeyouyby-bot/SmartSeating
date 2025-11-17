using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SmartSeating.Converters
{
    public sealed class NullOrEmptyToVisibilityConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            var hasContent = value is string text && !string.IsNullOrWhiteSpace(text);
            return hasContent ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
            => Binding.DoNothing;
    }
}
