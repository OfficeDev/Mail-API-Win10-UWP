using System;
using Windows.UI.Xaml.Data;

namespace MailClientWin10App.Converters
{
    class EmailDateToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            DateTimeOffset? dateVal = value as DateTimeOffset?;
            if (dateVal == null || !dateVal.HasValue)
                return value;

            var myDate = dateVal.Value.ToLocalTime();
            string retVal = string.Empty;
            if (myDate.Date == DateTime.Today)
            {
                retVal = myDate.ToString("h:mm tt");
            }
            else if (myDate.Date > DateTime.Today.AddDays(-6))
            {
                retVal = myDate.ToString("ddd h:mm tt");
            }
            else if (myDate.Year == DateTime.Today.Year)
            {
                retVal = myDate.Date.ToString("ddd M/dd");
            }
            else
            {
                retVal = myDate.Date.ToString("ddd M/dd/yy");
            }

            return retVal;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
}
