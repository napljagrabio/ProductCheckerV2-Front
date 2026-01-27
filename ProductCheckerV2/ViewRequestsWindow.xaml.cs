using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using ProductCheckerV2.Common;
using ProductCheckerV2.Database;
using ProductCheckerV2.Database.Models;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Windows.Media.Effects;

namespace ProductCheckerV2
{
    public partial class ViewRequestsPage : Page
    {
        private List<RequestViewModel> _allRequests = new List<RequestViewModel>();
        private ICollectionView _requestsView;
        private RequestViewModel _selectedRequest;
        private DispatcherTimer _autoRefreshTimer;
        private DateTime _lastRefreshTime;
        private bool _isRefreshing = false;
        private bool _rescanOnlyErrors = false;
        private string _selectedExportPath = "";
        private string _currentSearchText = string.Empty;

        public ViewRequestsPage()
        {
            InitializeComponent();
            Loaded += ViewRequestsPage_Loaded;
            Unloaded += ViewRequestsPage_Unloaded;
        }

        private void ViewRequestsPage_Loaded(object sender, RoutedEventArgs e)
        {
            InitializeAutoRefresh();
            LoadRequests();
        }

        private void ViewRequestsPage_Unloaded(object sender, RoutedEventArgs e)
        {
            StopAutoRefresh();
        }


        private void InitializeAutoRefresh()
        {
            // Set up auto-refresh timer for real-time updates (every 5 seconds)
            _autoRefreshTimer = new DispatcherTimer();
            _autoRefreshTimer.Interval = TimeSpan.FromSeconds(5);
            _autoRefreshTimer.Tick += AutoRefreshTimer_Tick;
            _autoRefreshTimer.Start();

            _lastRefreshTime = DateTime.Now;
        }

        private void AutoRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (_isRefreshing) return; // Prevent overlapping refreshes

            _isRefreshing = true;

            try
            {
                // Use Dispatcher to ensure UI updates happen on the UI thread
                Dispatcher.InvokeAsync(async () =>
                {
                    await RefreshDataAsync();
                    _lastRefreshTime = DateTime.Now;
                }, DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                // Log error but don't show to user for auto-refresh
                Console.WriteLine($"Auto-refresh error: {ex.Message}");
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        private async Task RefreshDataAsync()
        {
            try
            {
                using var context = new ProductCheckerDbContext();

                var requests = await context.Requests
                    .Where(r => r.RequestInfoId > 0 && r.DeletedAt == null)
                    .Select(r => new
                    {
                        r.Id,
                        r.Priority,
                        r.Status,
                        r.CreatedAt,
                        r.RequestInfoId,
                        RequestInfo = r.RequestInfo != null ? new
                        {
                            r.RequestInfo.User,
                            r.RequestInfo.FileName
                        } : null
                    })
                    .ToListAsync();

                var requestViewModels = new List<RequestViewModel>();

                foreach (var r in requests)
                {
                    int listingsCount = 0;
                    if (r.RequestInfoId > 0)
                    {
                        listingsCount = await context.ProductListings
                            .CountAsync(l => l.RequestInfoId == r.RequestInfoId);
                    }

                    requestViewModels.Add(new RequestViewModel
                    {
                        Id = r.Id,
                        RequestInfoId = r.RequestInfoId,
                        User = r.RequestInfo?.User ?? "Unknown",
                        FileName = r.RequestInfo?.FileName ?? "Unknown",
                        Status = r.Status,
                        CreatedAt = r.CreatedAt,
                        ListingsCount = listingsCount,
                        Priority = r.Priority,
                        IsHighPriority = r.Priority == 1
                    });
                }

                requestViewModels = requestViewModels.OrderByDescending(r => r.Id).ToList();

                // Check if there are any changes
                bool hasChanges = CheckForChanges(requestViewModels);
                bool searchActive = !string.IsNullOrWhiteSpace(SearchTextBox?.Text);

                if (hasChanges)
                {
                    // Update the requests list
                    UpdateRequestsList(requestViewModels);
                }

                // Always refresh the selected request/listings so the right panel stays current
                if (_selectedRequest != null)
                {
                    var selectedRequestInDb = requestViewModels.FirstOrDefault(r => r.Id == _selectedRequest.Id);
                    if (selectedRequestInDb != null)
                    {
                        // Update selected request status and listings count
                        _selectedRequest.Status = selectedRequestInDb.Status;
                        _selectedRequest.ListingsCount = selectedRequestInDb.ListingsCount;
                        _selectedRequest.StatusBrush = GetStatusBrush(_selectedRequest.Status);
                        _selectedRequest.Priority = selectedRequestInDb.Priority;
                        _selectedRequest.IsHighPriority = selectedRequestInDb.Priority == 1;

                        // Update UI for selected request
                        UpdateSelectedRequestDisplay(_selectedRequest);

                        // Refresh listings for the selected request on every auto-refresh tick
                        LoadListingsForRequest(_selectedRequest.Id);
                    }
                    else
                    {
                        ClearListings();
                        UpdateSelectedRequestDisplay(null);
                    }
                }

                if (hasChanges)
                {
                    if (searchActive)
                    {
                        UpdateListingSearchMatches();
                    }

                    // Refresh the filtered view
                    _requestsView?.Refresh();
                    UpdateRequestCountText();

                    // Show a subtle notification that data was refreshed (optional)
                    ShowRefreshNotification();
                }
                else if (searchActive)
                {
                    // Keep search results fresh even when request metadata hasn't changed.
                    UpdateListingSearchMatches();
                    _requestsView?.Refresh();
                    UpdateRequestCountText();
                }
            }
            catch (Exception ex)
            {
                // Don't show error messages for auto-refresh failures
                Console.WriteLine($"Auto-refresh error: {ex.Message}");
            }
        }

        private void LoadRequests()
        {
            try
            {
                using var context = new ProductCheckerDbContext();

                var requests = context.Requests
                    .Where(r => r.RequestInfoId > 0 && r.DeletedAt == null)
                    .Select(r => new
                    {
                        r.Id,
                        r.Priority,
                        r.Status,
                        r.CreatedAt,
                        r.RequestInfoId,
                        RequestInfo = r.RequestInfo != null ? new
                        {
                            r.RequestInfo.User,
                            r.RequestInfo.FileName
                        } : null
                    })
                    .ToList();

                _allRequests.Clear();

                int pendingRequests = 0;

                foreach (var r in requests)
                {
                    int listingsCount = 0;
                    if (r.RequestInfoId > 0)
                    {
                        listingsCount = context.ProductListings
                            .Count(l => l.RequestInfoId == r.RequestInfoId);
                    }

                    if (r.Status == RequestStatus.PENDING)
                    {
                        pendingRequests++;
                    }

                    var requestVM = new RequestViewModel
                    {
                        Id = r.Id,
                        RequestInfoId = r.RequestInfoId,
                        User = r.RequestInfo?.User ?? "Unknown",
                        FileName = r.RequestInfo?.FileName ?? "Unknown",
                        Status = r.Status,
                        CreatedAt = r.CreatedAt,
                        ListingsCount = listingsCount,
                        Priority = r.Priority,
                        IsHighPriority = r.Priority == 1
                    };

                    requestVM.StatusBrush = GetStatusBrush(requestVM.Status);
                    _allRequests.Add(requestVM);
                }

                _allRequests = _allRequests.OrderByDescending(r => r.Id).ToList();

                _requestsView = CollectionViewSource.GetDefaultView(_allRequests);
                _requestsView.Filter = RequestFilter;

                RequestsListBox.ItemsSource = _requestsView;
                if (!string.IsNullOrWhiteSpace(SearchTextBox?.Text))
                {
                    UpdateListingSearchMatches();
                    _requestsView.Refresh();
                    UpdateRequestCountText();
                }
                else
                {
                    RequestCountText.Text = $"({pendingRequests} pending)";
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_allRequests.Any())
                    {
                        if (RequestsListBox.ItemsSource != null && RequestsListBox.Items.Count > 0)
                        {
                            try
                            {
                                RequestsListBox.SelectedIndex = 0;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error selecting first request: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        UpdateSelectedRequestDisplay(null);
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);

                _lastRefreshTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading requests: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadListingsForRequest(int requestId)
        {
            try
            {
                using var context = new ProductCheckerDbContext();

                // Get the RequestInfoId for this request
                var request = context.Requests
                    .Where(r => r.Id == requestId)
                    .Select(r => new { r.RequestInfoId })
                    .FirstOrDefault();

                if (request == null || request.RequestInfoId == 0)
                {
                    ListingsDataGrid.ItemsSource = null;
                    ListingCountText.Text = "0 listings";
                    UpdateListingProgress(null);
                    return;
                }

                // Use raw SQL to avoid EF Core translation issues
                var connection = context.Database.GetDbConnection();
                try
                {
                    if (connection.State != System.Data.ConnectionState.Open)
                        connection.Open();

                    using var command = connection.CreateCommand();
                    command.CommandText = @"
                SELECT 
                    listing_id, 
                    case_number, 
                    platform, 
                    url, 
                    status, 
                    checked_date, 
                    error_detail, 
                    note 
                FROM product_checker_listings 
                WHERE request_info_id = @requestInfoId 
                ORDER BY id";

                    var param = command.CreateParameter();
                    param.ParameterName = "@requestInfoId";
                    param.Value = request.RequestInfoId;
                    command.Parameters.Add(param);

                    using var reader = command.ExecuteReader();
                    var listings = new List<ListingViewModel>();

                    while (reader.Read())
                    {
                        listings.Add(new ListingViewModel
                        {
                            ListingId = SafeGetString(reader, 0),
                            CaseNumber = SafeGetString(reader, 1),
                            Platform = SafeGetString(reader, 2),
                            Url = SafeGetString(reader, 3),
                            UrlStatus = SafeGetString(reader, 4),
                            CheckedDate = SafeGetString(reader, 5),
                            Notes = SafeGetString(reader, 7)
                        });
                    }

                    reader.Close();

                    ListingsDataGrid.ItemsSource = listings;
                    ListingCountText.Text = $"{listings.Count} listings";
                    UpdateListingProgress(listings);
                }
                finally
                {
                    if (connection.State == System.Data.ConnectionState.Open)
                        connection.Close();
                }
            }
            catch (Exception ex)
            {
                if (!_isRefreshing)
                {
                    MessageBox.Show($"Error loading listings: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                ListingsDataGrid.ItemsSource = null;
                ListingCountText.Text = "Error loading listings";
                UpdateListingProgress(null);
            }
        }

        private string SafeGetString(DbDataReader reader, int columnIndex)
        {
            if (reader.IsDBNull(columnIndex))
                return string.Empty;

            return reader.GetString(columnIndex);
        }

        private void UpdateListingProgress(IReadOnlyCollection<ListingViewModel> listings)
        {
            int total = listings?.Count ?? 0;
            int processed = 0;

            if (listings != null)
            {
                processed = listings.Count(IsListingProcessed);
                if (processed > total)
                {
                    processed = total;
                }
            }

            ListingProgressBar.Maximum = Math.Max(1, total);
            ListingProgressBar.Value = processed;
            ListingProgressText.Text = $"{processed}/{total}";
            if (processed == total)
            {
                ProgressName.Text = "Completed:";
            }
            else
            {
                ProgressName.Text = "Progress:";
            }
        }

        private static bool IsListingProcessed(ListingViewModel listing)
        {
            if (listing == null)
            {
                return false;
            }

            var status = listing.UrlStatus;
            if (string.IsNullOrWhiteSpace(status))
            {
                return false;
            }

            return !status.Equals("PENDING", StringComparison.OrdinalIgnoreCase);
        }

        private bool CheckForChanges(List<RequestViewModel> newRequests)
        {
            if (_allRequests.Count != newRequests.Count)
                return true;

            for (int i = 0; i < _allRequests.Count; i++)
            {
                var oldRequest = _allRequests[i];
                var newRequest = newRequests.FirstOrDefault(r => r.Id == oldRequest.Id);

                if (newRequest == null)
                    return true;

                if (oldRequest.Status != newRequest.Status ||
                    oldRequest.ListingsCount != newRequest.ListingsCount ||
                    oldRequest.Priority != newRequest.Priority)
                {
                    return true;
                }
            }

            // Check for new requests
            foreach (var newRequest in newRequests)
            {
                if (!_allRequests.Any(r => r.Id == newRequest.Id))
                    return true;
            }

            return false;
        }

        private void UpdateRequestsList(List<RequestViewModel> newRequests)
        {
            foreach (var existingRequest in _allRequests.ToList())
            {
                var updatedRequest = newRequests.FirstOrDefault(r => r.Id == existingRequest.Id);
                if (updatedRequest != null)
                {
                    existingRequest.Status = updatedRequest.Status;
                    existingRequest.ListingsCount = updatedRequest.ListingsCount;
                    existingRequest.StatusBrush = GetStatusBrush(existingRequest.Status);
                    existingRequest.RequestInfoId = updatedRequest.RequestInfoId;
                    existingRequest.Priority = updatedRequest.Priority;
                    existingRequest.IsHighPriority = updatedRequest.IsHighPriority;
                }
            }

            var pendingRequests = 0;
            var newRequestIds = newRequests.Select(r => r.Id).Except(_allRequests.Select(r => r.Id));
            foreach (var newId in newRequestIds)
            {
                var request = newRequests.First(r => r.Id == newId);
                var requestVM = new RequestViewModel
                {
                    Id = request.Id,
                    RequestInfoId = request.RequestInfoId,
                    User = request.User,
                    FileName = request.FileName,
                    Status = request.Status,
                    CreatedAt = request.CreatedAt,
                    ListingsCount = request.ListingsCount,
                    StatusBrush = GetStatusBrush(request.Status),
                    Priority = request.Priority,
                    IsHighPriority = request.Priority == 1
                };

                if (request.Status == RequestStatus.PENDING)
                {
                    pendingRequests++;
                }
                _allRequests.Add(requestVM);
            }

            _allRequests = _allRequests.OrderByDescending(r => r.Id).ToList();
            if (!string.IsNullOrWhiteSpace(SearchTextBox?.Text))
            {
                UpdateRequestCountText();
            }
            else
            {
                RequestCountText.Text = $"({pendingRequests} pending)";
            }
        }

        private bool IsFinalStatus(RequestStatus status)
        {
            return status == RequestStatus.SUCCESS ||
                   status == RequestStatus.FAILED ||
                   status == RequestStatus.COMPLETED_WITH_ISSUES;
        }

        private void ShowRefreshNotification()
        {
            // Optional: Add a subtle animation or indicator that data was refreshed
        }

        private bool RequestFilter(object item)
        {
            if (string.IsNullOrWhiteSpace(SearchTextBox.Text))
                return true;

            if (item is RequestViewModel request)
            {
                var searchText = SearchTextBox.Text.ToLower();
                var requestMatch =
                    request.Id.ToString().Contains(searchText) ||
                    request.User.ToLower().Contains(searchText) ||
                    request.FileName.ToLower().Contains(searchText) ||
                    request.Status.ToString().ToLower().Contains(searchText);

                var listingMatch = request.HasListingMatch && request.MatchCount > 0;
                return requestMatch || listingMatch;
            }

            return false;
        }

        private SolidColorBrush GetStatusBrush(RequestStatus status)
        {
            return status switch
            {
                RequestStatus.PENDING => new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                RequestStatus.PROCESSING => CreateBlinkingBrush(Color.FromRgb(66, 165, 245)),
                RequestStatus.SUCCESS => new SolidColorBrush(Color.FromRgb(102, 187, 106)),
                RequestStatus.FAILED => new SolidColorBrush(Color.FromRgb(239, 83, 80)),
                RequestStatus.COMPLETED_WITH_ISSUES => new SolidColorBrush(Color.FromRgb(255, 214, 170)),
                _ => new SolidColorBrush(Colors.Gray)
            };
        }

        private SolidColorBrush CreateBlinkingBrush(Color baseColor)
        {
            var brush = new SolidColorBrush(baseColor);

            // Create animation for blinking effect
            var opacityAnimation = new DoubleAnimation
            {
                From = 0.4,
                To = 1.0,
                Duration = TimeSpan.FromSeconds(0.8),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
            };

            // Apply animation to brush
            brush.BeginAnimation(SolidColorBrush.OpacityProperty, opacityAnimation);

            return brush;
        }

        private void RequestsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RequestsListBox.SelectedItem is RequestViewModel selectedRequest)
            {
                _selectedRequest = selectedRequest;
                LoadListingsForRequest(selectedRequest.Id);
                UpdateSelectedRequestDisplay(selectedRequest);
                ExportButton.IsEnabled = IsExportAllowed(selectedRequest.Status);
            }
            else
            {
                ClearListings();
                UpdateSelectedRequestDisplay(null);
                ExportButton.IsEnabled = false;
            }
        }

        private bool IsExportAllowed(RequestStatus status)
        {
            return status == RequestStatus.SUCCESS || status == RequestStatus.COMPLETED_WITH_ISSUES;
        }

        private void UpdateSelectedRequestDisplay(RequestViewModel request)
        {
            if (request == null)
            {
                SelectedRequestTitle.Text = "No request selected";
                SelectedRequestInfo.Text = "Select a request to view listings";
                SelectedRequestStatus.Text = "No request selected";
                ListingCountText.Text = "0 listings";
                UpdateListingProgress(null);
                ExportButton.Background = new SolidColorBrush(Color.FromRgb(102, 187, 106));
                ExportButton.IsEnabled = false;
                RescanButton.Visibility = Visibility.Collapsed;
                RescanButton.IsEnabled = false;
                PriorityToggleButton.IsChecked = false;
                PriorityToggleButton.IsEnabled = false;
                return;
            }

            SelectedRequestTitle.Text = request.FileName;
            SelectedRequestInfo.Text = $"RID: {request.Id} • User: {request.User} • Created: {request.CreatedAt:yyyy MMM d, h:mm tt}";
            SelectedRequestStatus.Text = $"Status: {request.Status}";
            PriorityToggleButton.IsChecked = request.IsHighPriority;
            PriorityToggleButton.IsEnabled = true;
            PriorityToggleButton.ToolTip = request.IsHighPriority ? "Click to remove highest priority" : "Click to set highest priority";

            if (IsExportAllowed(request.Status))
            {
                ExportButton.Background = new SolidColorBrush(Color.FromRgb(102, 187, 106));
                ExportButton.ToolTip = "Export listings to Excel";
                ExportButton.IsEnabled = true;
            }
            else
            {
                ExportButton.Background = new SolidColorBrush(Color.FromRgb(156, 163, 175));
                ExportButton.ToolTip = $"Export not available. Current status: {request.Status}. Only 'SUCCESS' or 'COMPLETED_WITH_ISSUES' statuses can be exported.";
                ExportButton.IsEnabled = false;
            }

            if (IsRescanAllowed(request.Status))
            {
                RescanButton.Visibility = Visibility.Visible;
                RescanButton.IsEnabled = true;
                RescanButton.ToolTip = "Create a rescan request for this file";
            }
            else
            {
                RescanButton.Visibility = Visibility.Collapsed;
                RescanButton.IsEnabled = false;
                RescanButton.ToolTip = null;
            }
        }

        private void PriorityToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedRequest == null)
                return;

            try
            {
                var setHighPriority = PriorityToggleButton.IsChecked == true;
                using var context = new ProductCheckerDbContext();
                var request = context.Requests.FirstOrDefault(r => r.Id == _selectedRequest.Id);
                if (request == null)
                {
                    ShowErrorMessage("Request not found.", "Priority Update Failed");
                    return;
                }

                request.Priority = setHighPriority ? 1 : 0;
                context.SaveChanges();

                _selectedRequest.Priority = request.Priority;
                _selectedRequest.IsHighPriority = request.Priority == 1;

                var listItem = _allRequests.FirstOrDefault(r => r.Id == _selectedRequest.Id);
                if (listItem != null)
                {
                    listItem.Priority = request.Priority;
                    listItem.IsHighPriority = request.Priority == 1;
                }

                _requestsView?.Refresh();
                UpdateSelectedRequestDisplay(_selectedRequest);
                ShowSuccessMessage(setHighPriority ? "Request marked as highest priority." : "Request priority cleared.", "Priority Updated");
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Error updating priority: {ex.Message}", "Priority Update Failed");
            }
        }

        private void ClearListings()
        {
            ListingsDataGrid.ItemsSource = null;
            ListingCountText.Text = "0 listings";
            UpdateListingProgress(null);
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateListingSearchMatches();
            _requestsView?.Refresh();
            UpdateRequestCountText();
        }

        private void UpdateListingSearchMatches()
        {
            _currentSearchText = SearchTextBox?.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(_currentSearchText))
            {
                foreach (var request in _allRequests)
                {
                    request.MatchCount = 0;
                    request.HasListingMatch = false;
                }
                return;
            }

            using var context = new ProductCheckerDbContext();
            var searchLower = _currentSearchText.ToLower();

            var matchCountsByRequestInfoId = context.ProductListings
                .Where(l => l.RequestInfoId > 0 && (
                    (l.ListingId ?? string.Empty).ToLower().Contains(searchLower) ||
                    (l.CaseNumber ?? string.Empty).ToLower().Contains(searchLower) ||
                    (l.Platform ?? string.Empty).ToLower().Contains(searchLower) ||
                    (l.Url ?? string.Empty).ToLower().Contains(searchLower) ||
                    (l.UrlStatus ?? string.Empty).ToLower().Contains(searchLower) ||
                    (l.ErrorDetail ?? string.Empty).ToLower().Contains(searchLower) ||
                    (l.Note ?? string.Empty).ToLower().Contains(searchLower)))
                .GroupBy(l => l.RequestInfoId)
                .Select(g => new { RequestInfoId = g.Key, Count = g.Count() })
                .ToList()
                .ToDictionary(x => x.RequestInfoId, x => x.Count);

            foreach (var request in _allRequests)
            {
                if (matchCountsByRequestInfoId.TryGetValue(request.RequestInfoId, out var count))
                {
                    request.MatchCount = count;
                    request.HasListingMatch = count > 0;
                }
                else
                {
                    request.MatchCount = 0;
                    request.HasListingMatch = false;
                }
            }
        }

        private void UpdateRequestCountText()
        {
            if (_requestsView == null)
            {
                return;
            }

            _currentSearchText = SearchTextBox?.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(_currentSearchText))
            {
                int pendingRequests = _allRequests.Count(r => r.Status == RequestStatus.PENDING);
                RequestCountText.Text = $"({pendingRequests} pending)";
                return;
            }

            int filteredCount = _requestsView.Cast<object>().Count();
            RequestCountText.Text = $"({filteredCount} requests)";
        }

        // ==================== MODAL METHODS ====================

        private void ShowModal(Border modal)
        {
            // Stop auto-refresh while modal is open
            if (_autoRefreshTimer != null && _autoRefreshTimer.IsEnabled)
            {
                _autoRefreshTimer.Stop();
            }

            ModalOverlay.Visibility = Visibility.Visible;
            modal.Visibility = Visibility.Visible;

            // Animate modal entrance
            var animation = new DoubleAnimation
            {
                From = 0.8,
                To = 1.0,
                Duration = TimeSpan.FromSeconds(0.2),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            modal.BeginAnimation(OpacityProperty, animation);
        }

        private void HideModal(Border modal)
        {
            var animation = new DoubleAnimation
            {
                From = 1.0,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.15),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
            };

            animation.Completed += (s, e) =>
            {
                modal.Visibility = Visibility.Collapsed;
                ModalOverlay.Visibility = Visibility.Collapsed;

                // Resume auto-refresh
                if (_autoRefreshTimer != null && !_autoRefreshTimer.IsEnabled)
                {
                    _autoRefreshTimer.Start();
                }
            };

            modal.BeginAnimation(OpacityProperty, animation);
        }

        // ==================== RESCAN MODAL METHODS ====================

        private void RescanButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedRequest == null || !IsRescanAllowed(_selectedRequest.Status))
                return;

            // Reset options
            _rescanOnlyErrors = false;
            SelectRescanOption(false);

            // Show modal
            ShowModal(RescanModal);
        }

        private void RescanAllOption_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _rescanOnlyErrors = false;
            SelectRescanOption(false);
        }

        private void RescanErrorsOption_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _rescanOnlyErrors = true;
            SelectRescanOption(true);
        }

        private void SelectRescanOption(bool errorsOnly)
        {
            if (errorsOnly)
            {
                RescanAllOption.BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240));
                RescanAllOption.Background = new SolidColorBrush(Color.FromRgb(247, 250, 252));

                RescanErrorsOption.BorderBrush = new SolidColorBrush(Color.FromRgb(245, 158, 11));
                RescanErrorsOption.Background = new SolidColorBrush(Color.FromRgb(254, 252, 232));
                RescanErrorsOption.Effect = new DropShadowEffect
                {
                    BlurRadius = 10,
                    ShadowDepth = 0,
                    Opacity = 0.2,
                    Color = Color.FromRgb(245, 158, 11)
                };
            }
            else
            {
                RescanErrorsOption.BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240));
                RescanErrorsOption.Background = new SolidColorBrush(Color.FromRgb(247, 250, 252));
                RescanErrorsOption.Effect = null;

                RescanAllOption.BorderBrush = new SolidColorBrush(Color.FromRgb(245, 158, 11));
                RescanAllOption.Background = new SolidColorBrush(Color.FromRgb(254, 252, 232));
                RescanAllOption.Effect = new DropShadowEffect
                {
                    BlurRadius = 10,
                    ShadowDepth = 0,
                    Opacity = 0.2,
                    Color = Color.FromRgb(245, 158, 11)
                };
            }

            ConfirmRescanButton.IsEnabled = true;
        }

        private void CancelRescanButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(RescanModal);
        }

        private void ConfirmRescanButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(RescanModal);
            CreateRescanRequest();
        }

        private bool IsRescanAllowed(RequestStatus status)
        {
            return status == RequestStatus.FAILED || status == RequestStatus.COMPLETED_WITH_ISSUES;
        }

        private void CreateRescanRequest()
        {
            if (_selectedRequest == null)
                return;

            try
            {
                using var context = new ProductCheckerDbContext();

                var rescanRequest = new Request
                {
                    RequestInfoId = _selectedRequest.RequestInfoId,
                    Status = RequestStatus.PENDING,
                    RescanInfoId = _rescanOnlyErrors ? 1 : 0,
                    CreatedAt = DateTime.Now
                };

                IQueryable<ProductListing> listingsQuery = context.ProductListings
                    .Where(l => l.RequestInfoId == _selectedRequest.RequestInfoId);

                var listings = listingsQuery.ToList();

                if (_rescanOnlyErrors)
                {
                    listings = listings
                        .Where(l =>
                            !string.IsNullOrWhiteSpace(l.UrlStatus) &&
                            !l.UrlStatus.Equals("Available", StringComparison.OrdinalIgnoreCase) &&
                            !l.UrlStatus.Equals("Not Available", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }

                foreach (var listing in listings)
                {
                    listing.UrlStatus = null;
                    listing.CheckedDate = null;
                    listing.ErrorDetail = null;
                    listing.Note = null;
                }

                context.Requests.Add(rescanRequest);

                var originalRequest = context.Requests.FirstOrDefault(r => r.Id == _selectedRequest.Id);
                if (originalRequest != null)
                {
                    originalRequest.DeletedAt = DateTime.Now;
                }

                context.SaveChanges();

                // Show success message
                ShowSuccessMessage($"Successfully queued for rescan.", "Rescan Queued");
                LoadRequests();
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Error creating rescan request: {ex.Message}", "Rescan Error");
            }
        }

        // ==================== EXPORT MODAL METHODS ====================

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedRequest == null)
                return;

            // Check if export is allowed based on status
            if (!IsExportAllowed(_selectedRequest.Status))
            {
                ShowErrorMessage($"Export is not available for requests with status: {_selectedRequest.Status}\n\n" +
                              "Only requests with 'SUCCESS' or 'COMPLETED_WITH_ISSUES' status can be exported.",
                              "Export Not Allowed");
                return;
            }

            // Get listings count
            int listingsCount = 0;
            if (ListingsDataGrid.ItemsSource is System.Collections.IEnumerable items)
            {
                listingsCount = items.Cast<object>().Count();
            }

            // Setup modal
            ExportFileName.Text = Path.GetFileNameWithoutExtension(_selectedRequest.FileName);
            ExportListingsCount.Text = listingsCount.ToString();
            ExportStatus.Text = _selectedRequest.Status.ToString();

            // Set default file path
            string baseFileName = Path.GetFileNameWithoutExtension(_selectedRequest.FileName);
            string formattedDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string defaultFileName = $"{baseFileName}-Result-{formattedDate}.xlsx";
            string defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            _selectedExportPath = Path.Combine(defaultFolder, defaultFileName);
            ExportFilePath.Text = _selectedExportPath;

            ConfirmExportButton.IsEnabled = true;

            // Show modal
            ShowModal(ExportModal);
        }

        private void BrowseExportPathButton_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = Path.GetFileName(_selectedExportPath),
                DefaultExt = ".xlsx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                _selectedExportPath = saveDialog.FileName;
                ExportFilePath.Text = _selectedExportPath;
            }
        }

        private void CancelExportButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(ExportModal);
        }

        private void ConfirmExportButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(ExportModal);
            ExportToExcel();
        }

        private void ExportToExcel()
        {
            try
            {
                // Create directory if it doesn't exist
                string directory = Path.GetDirectoryName(_selectedExportPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                ExportToExcel(_selectedRequest.Id, _selectedExportPath);

                // Show success message
                ShowSuccessMessage($"Listings exported successfully.", "Export Complete");
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Error exporting to Excel: {ex.Message}", "Export Error");
            }
        }

        private void ExportToExcel(int requestId, string filePath)
        {
            try
            {
                using var context = new ProductCheckerDbContext();

                // First get the RequestInfoId for this request
                var request = context.Requests
                    .Where(r => r.Id == requestId)
                    .Select(r => new { r.RequestInfoId, r.RequestInfo })
                    .FirstOrDefault();

                if (request == null)
                    throw new Exception($"Request #{requestId} not found");

                if (request.RequestInfoId == 0)
                    throw new Exception($"Request #{requestId} has invalid RequestInfoId (0)");

                // Use raw SQL to get listings
                var connection = context.Database.GetDbConnection();
                try
                {
                    if (connection.State != System.Data.ConnectionState.Open)
                        connection.Open();

                    using var command = connection.CreateCommand();
                    command.CommandText = @"
                SELECT 
                    listing_id, 
                    case_number, 
                    platform, 
                    url, 
                    status, 
                    checked_date, 
                    note 
                FROM product_checker_listings 
                WHERE request_info_id = @requestInfoId 
                ORDER BY id";

                    var param = command.CreateParameter();
                    param.ParameterName = "@requestInfoId";
                    param.Value = request.RequestInfoId;
                    command.Parameters.Add(param);

                    using var reader = command.ExecuteReader();
                    var listings = new List<ExcelListing>();

                    while (reader.Read())
                    {
                        listings.Add(new ExcelListing
                        {
                            ListingId = SafeGetString(reader, 0),
                            CaseNumber = SafeGetString(reader, 1),
                            Platform = SafeGetString(reader, 2),
                            Url = SafeGetString(reader, 3),
                            UrlStatus = SafeGetString(reader, 4),
                            CheckedDate = SafeGetString(reader, 5),
                            Notes = SafeGetString(reader, 6)
                        });
                    }

                    reader.Close();

                    if (listings.Count == 0)
                        throw new Exception($"No listings found for request #{requestId}");

                    using var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("Listings");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "Listing ID";
                    worksheet.Cell(1, 2).Value = "Case Number";
                    worksheet.Cell(1, 3).Value = "Platform";
                    worksheet.Cell(1, 4).Value = "Product URL";
                    worksheet.Cell(1, 5).Value = "URL Status";
                    worksheet.Cell(1, 6).Value = "Checked Date";
                    worksheet.Cell(1, 7).Value = "Notes";

                    // Add listing data
                    int row = 2;
                    foreach (var listing in listings)
                    {
                        worksheet.Cell(row, 1).Value = listing.ListingId;
                        worksheet.Cell(row, 2).Value = listing.CaseNumber;
                        worksheet.Cell(row, 3).Value = listing.Platform;
                        worksheet.Cell(row, 4).Value = listing.Url;
                        worksheet.Cell(row, 5).Value = listing.UrlStatus;
                        worksheet.Cell(row, 6).Value = listing.CheckedDate;
                        worksheet.Cell(row, 7).Value = listing.Notes;
                        row++;
                    }

                    var headerRange = worksheet.Range(1, 1, 1, 8);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(filePath);
                }
                finally
                {
                    if (connection.State == System.Data.ConnectionState.Open)
                        connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error exporting to Excel: {ex.Message}", ex);
            }
        }

        // ==================== UTILITY METHODS ====================

        private void ShowSuccessMessage(string message, string title)
        {
            var ownerWindow = Window.GetWindow(this);
            var messageBox = new Window
            {
                Title = title,
                Width = 420,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = ownerWindow != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen,
                Owner = ownerWindow,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = Brushes.Transparent,
                ShowInTaskbar = false
            };

            var chromeBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(20),
                Effect = new DropShadowEffect
                {
                    BlurRadius = 18,
                    ShadowDepth = 0,
                    Opacity = 0.18,
                    Color = Colors.Black
                }
            };

            var rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var headerGrid = new Grid
            {
                Background = new SolidColorBrush(Color.FromRgb(236, 253, 245)),
                Margin = new Thickness(-20, -20, -20, 15)
            };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var iconBadge = new Border
            {
                Width = 36,
                Height = 36,
                CornerRadius = new CornerRadius(18),
                Background = new SolidColorBrush(Color.FromRgb(16, 185, 129)),
                Margin = new Thickness(16, 12, 12, 12),
                Child = new TextBlock
                {
                    Text = "✓",
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            Grid.SetColumn(iconBadge, 0);
            headerGrid.Children.Add(iconBadge);

            var titleBlock = new TextBlock
            {
                Text = title,
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(16, 185, 129)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 0)
            };
            Grid.SetColumn(titleBlock, 1);
            headerGrid.Children.Add(titleBlock);

            var closeButton = new Button
            {
                Content = "✕",
                Width = 28,
                Height = 28,
                Margin = new Thickness(0, 10, 10, 10),
                Background = Brushes.Transparent,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            closeButton.Click += (s, e) => messageBox.Close();
            Grid.SetColumn(closeButton, 2);
            headerGrid.Children.Add(closeButton);

            rootGrid.Children.Add(headerGrid);

            var bodyPanel = new StackPanel
            {
                Margin = new Thickness(0, 0, 0, 10)
            };
            bodyPanel.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 14.5,
                Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105)),
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 5)
            });
            Grid.SetRow(bodyPanel, 1);
            rootGrid.Children.Add(bodyPanel);

            chromeBorder.Child = rootGrid;
            messageBox.Content = chromeBorder;

            var closeTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(1500)
            };
            closeTimer.Tick += (s, e) =>
            {
                closeTimer.Stop();
                messageBox.Close();
            };

            messageBox.Loaded += (s, e) => closeTimer.Start();
            messageBox.Show();
        }

        private void ShowErrorMessage(string message, string title)
        {
            // Create a custom message box window
            var ownerWindow = Window.GetWindow(this);
            var messageBox = new Window
            {
                Title = title,
                Width = 450,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = ownerWindow != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen,
                Owner = ownerWindow,
                WindowStyle = WindowStyle.ToolWindow,
                Background = new SolidColorBrush(Colors.White),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1)
            };

            var contentPanel = new StackPanel
            {
                Margin = new Thickness(20)
            };

            // Header
            var headerPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 15)
            };
            headerPanel.Children.Add(new TextBlock
            {
                Text = "❌",
                FontSize = 24,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerPanel.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(239, 83, 80))
            });
            contentPanel.Children.Add(headerPanel);

            // Message
            contentPanel.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 14,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 20)
            });

            // OK Button
            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                HorizontalAlignment = HorizontalAlignment.Center,
                Background = new SolidColorBrush(Color.FromRgb(239, 83, 80)),
                Foreground = Brushes.White,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            okButton.Click += (s, e) => { messageBox.Close(); };
            contentPanel.Children.Add(okButton);

            messageBox.Content = contentPanel;
            messageBox.ShowDialog();
        }

        private void StopAutoRefresh()
        {
            // Stop the auto-refresh timer when window is closing
            if (_autoRefreshTimer != null)
            {
                _autoRefreshTimer.Stop();
                _autoRefreshTimer.Tick -= AutoRefreshTimer_Tick;
                _autoRefreshTimer = null;
            }
        }

        // Helper class for Excel export
        private class ExcelListing
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

    // View Models
    public class RequestViewModel : INotifyPropertyChanged
    {
        public int Id { get; set; }
        public int RequestInfoId { get; set; }
        public string User { get; set; }
        public string FileName { get; set; }
        public RequestStatus Status { get; set; }
        public DateTime CreatedAt { get; set; }
        public int ListingsCount { get; set; }
        public SolidColorBrush StatusBrush { get; set; }
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
