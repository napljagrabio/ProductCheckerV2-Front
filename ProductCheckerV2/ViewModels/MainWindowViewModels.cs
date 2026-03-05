using System.ComponentModel;
using System.Windows.Media;

namespace ProductCheckerV2
{
    internal class ProcessingResult
    {
        public bool Success { get; set; }
        public long RequestId { get; set; }
        public int RecordsProcessed { get; set; }
        public string Message { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class FilterOption : INotifyPropertyChanged
    {
        public int Id { get; set; }
        public string Display { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public class SelectedFilterItem
    {
        public string Category { get; set; }
        public string Display { get; set; }
        public FilterOption Option { get; set; }
    }

    public class FileInfoItem
    {
        public string Text { get; set; }
        public Brush Color { get; set; }
    }

    public class UploadedProductData
    {
        public string ListingId { get; set; }
        public string CaseNumber { get; set; }
        public string ProductUrl { get; set; }
        public string Platform { get; set; }
    }
}

