using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using ProductCheckerV2.Database.Models;

namespace ProductCheckerV2.Common
{
    public class StatusColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is RequestStatus status)
            {
                return status switch
                {
                    RequestStatus.PENDING => Brushes.Orange,
                    RequestStatus.PROCESSING => Brushes.Blue,
                    RequestStatus.SUCCESS => Brushes.Green,
                    RequestStatus.FAILED => Brushes.Red,
                    _ => Brushes.Gray
                };
            }
            return Brushes.Gray;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
