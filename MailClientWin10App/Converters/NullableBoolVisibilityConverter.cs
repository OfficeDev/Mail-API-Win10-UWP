    using System;
    using Windows.UI.Xaml;
    using Windows.UI.Xaml.Data;

    namespace MailClientWin10App.Converters
    {
        class NullableBoolVisibilityConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, string language)
            {
                bool? b = value as bool?;
                if (b == null || !b.HasValue || !b.Value)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
                //if (b.Value) return Visibility.Visible;
                //else return Visibility.Collapsed;
            }

            public object ConvertBack(object value, Type targetType, object parameter, string language)
            {
                throw new NotImplementedException();
            }
        }
    }