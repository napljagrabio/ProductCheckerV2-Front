using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ProductCheckerV2.Controls
{
    public partial class PaginationControl : UserControl
    {
        public static readonly DependencyProperty CurrentPageProperty =
            DependencyProperty.Register(
                nameof(CurrentPage),
                typeof(int),
                typeof(PaginationControl),
                new PropertyMetadata(1, OnPaginationPropertyChanged));

        public static readonly DependencyProperty TotalItemsProperty =
            DependencyProperty.Register(
                nameof(TotalItems),
                typeof(int),
                typeof(PaginationControl),
                new PropertyMetadata(0, OnPaginationPropertyChanged));

        public static readonly DependencyProperty PageSizeProperty =
            DependencyProperty.Register(
                nameof(PageSize),
                typeof(int),
                typeof(PaginationControl),
                new PropertyMetadata(10, OnPaginationPropertyChanged));

        public static readonly DependencyProperty MaxVisiblePagesProperty =
            DependencyProperty.Register(
                nameof(MaxVisiblePages),
                typeof(int),
                typeof(PaginationControl),
                new PropertyMetadata(5, OnPaginationPropertyChanged));

        private bool _isUpdatingUi;

        public event EventHandler<PaginationChangedEventArgs>? PaginationChanged;

        public int CurrentPage
        {
            get => (int)GetValue(CurrentPageProperty);
            set => SetValue(CurrentPageProperty, value);
        }

        public int TotalItems
        {
            get => (int)GetValue(TotalItemsProperty);
            set => SetValue(TotalItemsProperty, value);
        }

        public int PageSize
        {
            get => (int)GetValue(PageSizeProperty);
            set => SetValue(PageSizeProperty, value);
        }

        public int MaxVisiblePages
        {
            get => (int)GetValue(MaxVisiblePagesProperty);
            set => SetValue(MaxVisiblePagesProperty, value);
        }

        public PaginationControl()
        {
            InitializeComponent();
            Loaded += PaginationControl_Loaded;
        }

        private int TotalPages => TotalItems <= 0
            ? 1
            : (int)Math.Ceiling(TotalItems / (double)Math.Max(1, PageSize));

        private static void OnPaginationPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is PaginationControl control)
            {
                control.UpdatePaginationVisuals();
            }
        }

        private void PaginationControl_Loaded(object sender, RoutedEventArgs e)
        {
            UpdatePaginationVisuals();
        }

        private void UpdatePaginationVisuals()
        {
            if (!IsLoaded || _isUpdatingUi)
            {
                return;
            }

            _isUpdatingUi = true;
            try
            {
                if (PageSize < 1)
                {
                    SetCurrentValue(PageSizeProperty, 1);
                }

                if (MaxVisiblePages < 1)
                {
                    SetCurrentValue(MaxVisiblePagesProperty, 1);
                }

                var totalPages = TotalPages;
                var clampedPage = Math.Max(1, Math.Min(CurrentPage, totalPages));
                if (CurrentPage != clampedPage)
                {
                    SetCurrentValue(CurrentPageProperty, clampedPage);
                }

                var hasItems = TotalItems > 0;
                var startItem = hasItems ? ((clampedPage - 1) * PageSize) + 1 : 0;
                var endItem = hasItems ? Math.Min(clampedPage * PageSize, TotalItems) : 0;

                ResultsSummaryText.Text = $"Results: {startItem} - {endItem} of {TotalItems}";
                PrevPageButton.IsEnabled = hasItems && clampedPage > 1;
                NextPageButton.IsEnabled = hasItems && clampedPage < totalPages;

                SyncPageSizeSelection();
                RenderPageButtons(totalPages, clampedPage);
            }
            finally
            {
                _isUpdatingUi = false;
            }
        }

        private void SyncPageSizeSelection()
        {
            var selected = PageSizeComboBox.Items
                .OfType<ComboBoxItem>()
                .FirstOrDefault(item => int.TryParse(item.Content?.ToString(), out var value) && value == PageSize);

            if (selected == null && PageSizeComboBox.Items.Count > 0)
            {
                selected = (ComboBoxItem)PageSizeComboBox.Items[0];
                if (int.TryParse(selected.Content?.ToString(), out var selectedSize))
                {
                    SetCurrentValue(PageSizeProperty, selectedSize);
                }
            }

            if (selected != null && !ReferenceEquals(PageSizeComboBox.SelectedItem, selected))
            {
                PageSizeComboBox.SelectedItem = selected;
            }
        }

        private void RenderPageButtons(int totalPages, int currentPage)
        {
            PageNumberPanel.Children.Clear();

            foreach (var pageNumber in BuildPaginationSlots(totalPages, currentPage, MaxVisiblePages))
            {
                var hasData = pageNumber <= totalPages;
                var pageButton = new Button
                {
                    Content = pageNumber.ToString(),
                    Tag = pageNumber,
                    IsEnabled = hasData,
                    Style = (Style)FindResource(pageNumber == currentPage
                        ? "PaginationActiveButtonStyle"
                        : "PaginationButtonStyle")
                };

                if (hasData)
                {
                    pageButton.Click += PageNumberButton_Click;
                }

                PageNumberPanel.Children.Add(pageButton);
            }
        }

        private static List<int> BuildPaginationSlots(int totalPages, int currentPage, int visibleSlots)
        {
            var slots = new List<int>();
            if (visibleSlots <= 0)
            {
                return slots;
            }

            totalPages = Math.Max(1, totalPages);
            currentPage = Math.Max(1, Math.Min(currentPage, totalPages));

            var startPage = 1;
            if (totalPages >= visibleSlots)
            {
                var halfWindow = visibleSlots / 2;
                startPage = currentPage - halfWindow;
                startPage = Math.Max(1, startPage);

                var maxStart = totalPages - visibleSlots + 1;
                startPage = Math.Min(startPage, maxStart);
            }

            for (var page = startPage; page < startPage + visibleSlots; page++)
            {
                slots.Add(page);
            }

            return slots;
        }

        private void PageNumberButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is int pageNumber && pageNumber != CurrentPage)
            {
                CurrentPage = pageNumber;
                RaisePaginationChanged();
            }
        }

        private void PrevPageButton_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentPage > 1)
            {
                CurrentPage--;
                RaisePaginationChanged();
            }
        }

        private void NextPageButton_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentPage < TotalPages)
            {
                CurrentPage++;
                RaisePaginationChanged();
            }
        }

        private void PageSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingUi)
            {
                return;
            }

            if (PageSizeComboBox.SelectedItem is ComboBoxItem selectedItem &&
                int.TryParse(selectedItem.Content?.ToString(), out var selectedSize) &&
                selectedSize != PageSize)
            {
                PageSize = selectedSize;
                CurrentPage = 1;
                RaisePaginationChanged();
            }
        }

        private void RaisePaginationChanged()
        {
            PaginationChanged?.Invoke(this, new PaginationChangedEventArgs(CurrentPage, PageSize));
        }
    }

    public class PaginationChangedEventArgs : EventArgs
    {
        public PaginationChangedEventArgs(int currentPage, int pageSize)
        {
            CurrentPage = currentPage;
            PageSize = pageSize;
        }

        public int CurrentPage { get; }

        public int PageSize { get; }
    }
}
