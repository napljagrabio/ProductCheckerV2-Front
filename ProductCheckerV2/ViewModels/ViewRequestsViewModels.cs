using ProductCheckerV2.Database.Models;
using System;
using System.ComponentModel;
using System.Windows.Media;

namespace ProductCheckerV2
{
    internal class ExcelListing
    {
        public string ListingId { get; set; }
        public string CaseNumber { get; set; }
        public string Platform { get; set; }
        public string Url { get; set; }
        public string UrlStatus { get; set; }
        public string CheckedDate { get; set; }
        public string Notes { get; set; }
    }

    public class RequestViewModel : INotifyPropertyChanged
    {
        public long Id { get; set; }
        public long RequestInfoId { get; set; }
        public string User { get; set; }
        public string FileName { get; set; }
        public string Environment { get; set; } = "Stage";
        public RequestStatus Status { get; set; }
        public DateTime CreatedAt { get; set; }
        public int ListingsCount { get; set; }
        public SolidColorBrush StatusBrush { get; set; }
        public SolidColorBrush EnvironmentBrush { get; set; } = new SolidColorBrush(Color.FromRgb(245, 158, 11));
        public int Priority { get; set; }
        public bool IsHighPriority { get; set; }

        private int _matchCount;
        public int MatchCount
        {
            get => _matchCount;
            set
            {
                if (_matchCount == value)
                {
                    return;
                }

                _matchCount = value;
                OnPropertyChanged(nameof(MatchCount));
            }
        }

        private bool _hasListingMatch;
        public bool HasListingMatch
        {
            get => _hasListingMatch;
            set
            {
                if (_hasListingMatch == value)
                {
                    return;
                }

                _hasListingMatch = value;
                OnPropertyChanged(nameof(HasListingMatch));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ListingViewModel
    {
        public string ListingId { get; set; }
        public string CaseNumber { get; set; }
        public string Platform { get; set; }
        public string Url { get; set; }
        public string UrlStatus { get; set; }
        public string CheckedDate { get; set; }
        public string Notes { get; set; }
    }
}
