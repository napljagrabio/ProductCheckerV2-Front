using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using ProductCheckerV2.Common;
using ProductCheckerV2.Database;
using ProductCheckerV2.Database.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Threading;

namespace ProductCheckerV2
{
    public partial class ViewRequestsPage : Page
    {
        private int _pageSize = 10;
        private List<RequestViewModel> _allRequests = new List<RequestViewModel>();
        private ICollectionView _requestsView;
        private RequestViewModel _selectedRequest;
        private DispatcherTimer _autoRefreshTimer;
        private DateTime _lastRefreshTime;
        private bool _isRefreshing = false;
        private bool _rescanOnlyErrors = false;
        private string _selectedExportPath = "";
        private string _currentSearchText = string.Empty;
        private int _currentPage = 1;
        private int _totalPages = 1;
        private int _totalRequests = 0;
        private int _pendingRequests = 0;
        private bool _hasShownConnectionIssueModal = false;
        private bool _hasShownCacheFallbackModal = false;
        private bool _hasStartedInitialLoad = false;
        private readonly DispatcherTimer _searchDebounceTimer;
        private const int SearchDebounceMilliseconds = 300;
        private const int PageTransitionMilliseconds = 120;
        private const int BackgroundRefreshIntervalSeconds = 5;
        private const int DataQueryTimeoutSeconds = 12;
        private readonly string _cacheFilePath;
        private readonly object _cacheLock = new();
        private readonly SemaphoreSlim _requestsRefreshLock = new(1, 1);
        private readonly SemaphoreSlim _listingsRefreshLock = new(1, 1);
        private ViewRequestsCacheData _cacheData = new();
        private bool _isRequestsLoading = false;
        private bool _isListingsLoading = false;
        private bool _suppressNextListingSkeleton = false;
        private bool _suppressNextListingReload = false;
        private bool _isRestoringRequestSelection = false;
        private bool _isConnectionAvailable = true;

        public ViewRequestsPage()
        {
            InitializeComponent();
            Loaded += ViewRequestsPage_Loaded;
            Unloaded += ViewRequestsPage_Unloaded;
            _cacheFilePath = GetCacheFilePath();
            LoadCacheFromDisk();

            _searchDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(SearchDebounceMilliseconds)
            };
            _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;
        }

        private void ViewRequestsPage_Loaded(object sender, RoutedEventArgs e)
        {
            if (_hasStartedInitialLoad)
            {
                return;
            }

            _hasStartedInitialLoad = true;
            InitializeAutoRefresh();
            Dispatcher.BeginInvoke(new Action(LoadRequests), DispatcherPriority.ContextIdle);
        }

        private void ViewRequestsPage_Unloaded(object sender, RoutedEventArgs e)
        {
            StopAutoRefresh();
            _searchDebounceTimer.Stop();
            _searchDebounceTimer.Tick -= SearchDebounceTimer_Tick;
        }


        private void InitializeAutoRefresh()
        {
            _autoRefreshTimer = new DispatcherTimer();
            _autoRefreshTimer.Interval = TimeSpan.FromSeconds(BackgroundRefreshIntervalSeconds);
            _autoRefreshTimer.Tick += AutoRefreshTimer_Tick;
            _autoRefreshTimer.Start();

            _lastRefreshTime = DateTime.Now;
        }

        private async void AutoRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (_isRefreshing) return; // Prevent overlapping refreshes

            _isRefreshing = true;

            try
            {
                await RefreshDataAsync();
                _lastRefreshTime = DateTime.Now;
            }
            catch (Exception ex)
            {
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
                var requestsUpdated = await RefreshRequestsCacheAsync(showErrors: false);
                if (requestsUpdated)
                {
                    TryLoadRequestsFromCache(preserveSelection: true);
                }

                if (_selectedRequest != null)
                {
                    var selectedRequestId = _selectedRequest.Id;
                    var listingsUpdated = await RefreshListingsCacheAsync(selectedRequestId, showErrors: false);
                    if (listingsUpdated && _selectedRequest?.Id == selectedRequestId)
                    {
                        TryLoadListingsIntoUiFromCache(selectedRequestId, showCachedLabel: false);
                    }
                }

                _lastRefreshTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Auto-refresh error: {ex.Message}");
            }
        }

        private async void LoadRequests()
        {
            await LoadRequestsPageAsync(preserveSelection: false, showErrors: true);
        }

        private async Task LoadRequestsPageAsync(bool preserveSelection, bool showErrors, bool showSkeleton = true)
        {
            bool loadedFromCache = false;

            if (showSkeleton)
            {
                SetRequestsLoadingState(true);
            }

            try
            {
                loadedFromCache = TryLoadRequestsFromCache(preserveSelection);
                var refreshSucceeded = _isConnectionAvailable
                    ? await RefreshRequestsCacheAsync(showErrors)
                    : false;

                if (refreshSucceeded)
                {
                    TryLoadRequestsFromCache(preserveSelection);
                }
                else if (loadedFromCache && !_hasShownCacheFallbackModal)
                {
                    _hasShownCacheFallbackModal = true;
                    ModalDialogService.ShowConnectionLostBanner(Window.GetWindow(this));
                }
                else if (!loadedFromCache)
                {
                    _totalRequests = 0;
                    _pendingRequests = 0;
                    _totalPages = 1;
                    _currentPage = 1;
                    _allRequests = new List<RequestViewModel>();
                    _requestsView = CollectionViewSource.GetDefaultView(_allRequests);
                    _requestsView.Filter = RequestFilter;
                    RequestsListBox.ItemsSource = _requestsView;
                    _requestsView.Refresh();
                    UpdateRequestCountText();
                    UpdatePaginationControls();
                    ApplySelectionAfterLoad(preserveSelection);
                }
            }
            finally
            {
                if (showSkeleton)
                {
                    SetRequestsLoadingState(false);
                }
            }
        }

        private async Task<bool> RefreshRequestsCacheAsync(bool showErrors)
        {
            await _requestsRefreshLock.WaitAsync();

            try
            {
                var snapshot = await FetchRequestsSnapshotAsync();

                lock (_cacheLock)
                {
                    _cacheData.Requests = snapshot
                        .OrderByDescending(r => r.Id)
                        .Take(500)
                        .ToList();
                }

                SaveCacheToDisk();
                _hasShownConnectionIssueModal = false;
                _hasShownCacheFallbackModal = false;
                MarkConnectionRestored();
                return true;
            }
            catch (Exception ex)
            {
                if (IsDatabaseConnectionIssue(ex))
                {
                    MarkConnectionLost();
                }

                if (showErrors)
                {
                    //NOTE: Nothing to do

                    //ModalDialogService.Show(
                    //    $"Error loading requests: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    //    "Error",
                    //    MessageBoxButton.OK,
                    //    MessageBoxImage.Error);
                }
                else
                {
                    if (!_hasShownConnectionIssueModal && IsDatabaseConnectionIssue(ex))
                    {
                        _hasShownConnectionIssueModal = true;
                    }

                    Console.WriteLine($"Error loading requests: {ex.Message}");
                }

                return false;
            }
            finally
            {
                _requestsRefreshLock.Release();
            }
        }

        private async Task<List<CachedRequest>> FetchRequestsSnapshotAsync()
        {
            return await WithTimeoutAsync(
                _ => Task.Run(() =>
                {
                    using var context = new ProductCheckerDbContext();

                    var requests = context.Requests
                        .AsNoTracking()
                        .Where(r => r.RequestInfoId > 0 && r.DeletedAt == null)
                        .OrderByDescending(r => r.Id)
                        .Take(500)
                        .Select(r => new
                        {
                            r.Id,
                            r.RequestInfoId,
                            r.Priority,
                            r.Status,
                            r.CreatedAt,
                            User = r.RequestInfo != null ? r.RequestInfo.User : null,
                            FileName = r.RequestInfo != null ? r.RequestInfo.FileName : null,
                            Environment = r.RequestInfo != null ? r.RequestInfo.Environment : null
                        })
                        .ToList();

                    var requestInfoIds = requests
                        .Select(r => r.RequestInfoId)
                        .Where(id => id > 0)
                        .Distinct()
                        .ToList();

                    var listingCountsByRequestInfoId = requestInfoIds.Count == 0
                        ? new Dictionary<long, int>()
                        : context.ProductListings
                            .AsNoTracking()
                            .Where(l => requestInfoIds.Contains(l.RequestInfoId))
                            .GroupBy(l => l.RequestInfoId)
                            .Select(g => new { RequestInfoId = g.Key, Count = g.Count() })
                            .ToDictionary(x => x.RequestInfoId, x => x.Count);

                    return requests.Select(r =>
                    {
                        listingCountsByRequestInfoId.TryGetValue(r.RequestInfoId, out var listingsCount);
                        return new CachedRequest
                        {
                            Id = r.Id,
                            RequestInfoId = r.RequestInfoId,
                            User = r.User ?? "Unknown",
                            FileName = r.FileName ?? "Unknown",
                            Environment = NormalizeEnvironment(r.Environment),
                            Status = r.Status.ToString(),
                            CreatedAt = r.CreatedAt,
                            ListingsCount = listingsCount,
                            Priority = r.Priority
                        };
                    }).ToList();
                }),
                DataQueryTimeoutSeconds);
        }

        private static bool IsDatabaseConnectionIssue(Exception ex)
        {
            var message = $"{ex.Message} {ex.InnerException?.Message}".ToLowerInvariant();

            return message.Contains("unable to connect") ||
                   message.Contains("can't connect") ||
                   message.Contains("cannot connect") ||
                   message.Contains("access denied") ||
                   message.Contains("authentication") ||
                   message.Contains("password") ||
                   message.Contains("connection");
        }

        private void MarkConnectionLost()
        {
            _isConnectionAvailable = false;
            ModalDialogService.ShowConnectionLostBanner(Window.GetWindow(this));
        }

        private void MarkConnectionRestored()
        {
            _isConnectionAvailable = true;
            ModalDialogService.ShowConnectionRestoredBanner(Window.GetWindow(this));
        }

        private void ApplySelectionAfterLoad(bool preserveSelection)
        {
            _isRestoringRequestSelection = true;

            try
            {
                if (!_allRequests.Any())
                {
                    _selectedRequest = null;
                    RequestsListBox.SelectedItem = null;
                    ClearListings();
                    UpdateSelectedRequestDisplay(null);
                    return;
                }

                if (preserveSelection && _selectedRequest != null)
                {
                    var match = _allRequests.FirstOrDefault(r => r.Id == _selectedRequest.Id);
                    if (match != null)
                    {
                        _suppressNextListingReload = true;
                        _suppressNextListingSkeleton = true;
                        RequestsListBox.SelectedItem = match;
                        return;
                    }
                }

                RequestsListBox.SelectedIndex = 0;
            }
            finally
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _isRestoringRequestSelection = false;
                }), DispatcherPriority.ContextIdle);
            }
        }

        private void SetRequestsLoadingState(bool isLoading)
        {
            if (_isRequestsLoading == isLoading)
            {
                return;
            }

            _isRequestsLoading = isLoading;

            if (RequestsSkeletonOverlay != null)
            {
                RequestsSkeletonOverlay.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            }

            if (RequestsListBox != null)
            {
                RequestsListBox.Visibility = isLoading ? Visibility.Hidden : Visibility.Visible;
            }

            if (PaginationContainer != null)
            {
                PaginationContainer.Opacity = isLoading ? 0.55 : 1.0;
                PaginationContainer.IsEnabled = !isLoading;
            }
        }

        private void UpdatePaginationControls()
        {
            if (PageInfoText != null)
            {
                PageInfoText.Text = $"Page {_currentPage} of {_totalPages}";
            }

            if (PageRangeText != null)
            {
                if (_totalRequests == 0)
                {
                    PageRangeText.Text = "0-0 of 0";
                }
                else
                {
                    int start = ((_currentPage - 1) * _pageSize) + 1;
                    int end = Math.Min(_currentPage * _pageSize, _totalRequests);
                    PageRangeText.Text = $"{start}-{end} of {_totalRequests}";
                }
            }

            if (PrevPageButton != null)
            {
                PrevPageButton.IsEnabled = _currentPage > 1;
            }

            if (NextPageButton != null)
            {
                NextPageButton.IsEnabled = _currentPage < _totalPages;
            }

            if (FirstPageButton != null)
            {
                FirstPageButton.IsEnabled = _currentPage > 1;
            }

            if (LastPageButton != null)
            {
                LastPageButton.IsEnabled = _currentPage < _totalPages;
            }

            if (GoToPageButton != null)
            {
                GoToPageButton.IsEnabled = _totalRequests > 0;
            }

            if (PageJumpTextBox != null && !PageJumpTextBox.IsKeyboardFocusWithin)
            {
                PageJumpTextBox.Text = _currentPage.ToString(CultureInfo.InvariantCulture);
            }
        }

        private async Task NavigateToPageAsync(int targetPage)
        {
            if (targetPage < 1 || targetPage > _totalPages || targetPage == _currentPage)
            {
                return;
            }

            _currentPage = targetPage;
            await RunPageTransitionAsync();
            await LoadRequestsPageAsync(preserveSelection: true, showErrors: false);
        }

        private async Task RunPageTransitionAsync()
        {
            if (RequestsListBox == null)
            {
                return;
            }

            var fadeOut = new DoubleAnimation
            {
                To = 0.35,
                Duration = TimeSpan.FromMilliseconds(PageTransitionMilliseconds),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            await BeginAnimationAsync(RequestsListBox, UIElement.OpacityProperty, fadeOut);

            var fadeIn = new DoubleAnimation
            {
                To = 1.0,
                Duration = TimeSpan.FromMilliseconds(PageTransitionMilliseconds),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            await BeginAnimationAsync(RequestsListBox, UIElement.OpacityProperty, fadeIn);
        }

        private static Task BeginAnimationAsync(UIElement target, DependencyProperty property, AnimationTimeline animation)
        {
            var tcs = new TaskCompletionSource<bool>();

            if (animation == null)
            {
                tcs.SetResult(true);
                return tcs.Task;
            }

            void OnCompleted(object? sender, EventArgs e)
            {
                animation.Completed -= OnCompleted;
                tcs.TrySetResult(true);
            }

            animation.Completed += OnCompleted;
            Storyboard.SetTarget(animation, target);
            Storyboard.SetTargetProperty(animation, new PropertyPath(property));

            var storyboard = new Storyboard();
            storyboard.Children.Add(animation);
            storyboard.Begin();

            return tcs.Task;
        }

        private async Task LoadListingsForRequestAsync(long requestId, bool showSkeleton = true)
        {
            bool loadedFromCache = false;

            if (showSkeleton)
            {
                SetListingsLoadingState(true);
            }

            try
            {
                loadedFromCache = TryLoadListingsIntoUiFromCache(requestId, showCachedLabel: true);
                var refreshed = _isConnectionAvailable
                    ? await RefreshListingsCacheAsync(requestId, showErrors: !_isRefreshing)
                    : false;

                if (refreshed)
                {
                    TryLoadListingsIntoUiFromCache(requestId, showCachedLabel: false);
                }
                else if (!loadedFromCache)
                {
                    ListingsDataGrid.ItemsSource = null;
                    ListingCountText.Text = "Error loading listings";
                    UpdateListingProgress(null);
                }
            }
            finally
            {
                if (showSkeleton)
                {
                    SetListingsLoadingState(false);
                }
            }
        }

        private async Task<bool> RefreshListingsCacheAsync(long requestId, bool showErrors)
        {
            await _listingsRefreshLock.WaitAsync();

            try
            {
                var snapshot = await FetchListingsSnapshotAsync(requestId);
                if (snapshot == null)
                {
                    return false;
                }

                lock (_cacheLock)
                {
                    _cacheData.ListingsByRequestId[requestId] = snapshot;
                }

                SaveCacheToDisk();
                MarkConnectionRestored();
                return true;
            }
            catch (Exception ex)
            {
                if (IsDatabaseConnectionIssue(ex))
                {
                    MarkConnectionLost();
                }

                if (showErrors)
                {
                    //NOTE: Nothing to do

                    //ModalDialogService.Show(
                    //    $"Error loading listings: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    //    "Error",
                    //    MessageBoxButton.OK,
                    //    MessageBoxImage.Error);
                }
                else
                {
                    Console.WriteLine($"Error loading listings: {ex.Message}");
                }

                return false;
            }
            finally
            {
                _listingsRefreshLock.Release();
            }
        }

        private async Task<List<CachedListing>> FetchListingsSnapshotAsync(long requestId)
        {
            return await WithTimeoutAsync(
                _ => Task.Run(() =>
                {
                    using var context = new ProductCheckerDbContext();

                    var requestInfoId = context.Requests
                        .AsNoTracking()
                        .Where(r => r.Id == requestId)
                        .Select(r => r.RequestInfoId)
                        .FirstOrDefault();

                    if (requestInfoId <= 0)
                    {
                        return new List<CachedListing>();
                    }

                    return context.ProductListings
                        .AsNoTracking()
                        .Where(l => l.RequestInfoId == requestInfoId)
                        .OrderBy(l => l.Id)
                        .Select(l => new CachedListing
                        {
                            ListingId = l.ListingId.ToString(),
                            CaseNumber = l.CaseNumber ?? string.Empty,
                            Platform = l.Platform ?? string.Empty,
                            Url = l.Url ?? string.Empty,
                            UrlStatus = l.UrlStatus ?? string.Empty,
                            CheckedDate = l.CheckedDate ?? string.Empty,
                            Notes = l.Note ?? string.Empty
                        })
                        .ToList();
                }),
                DataQueryTimeoutSeconds);
        }

        private bool TryLoadListingsIntoUiFromCache(long requestId, bool showCachedLabel)
        {
            if (!TryLoadListingsFromCache(requestId, out var listings))
            {
                return false;
            }

            if (ShouldReplaceListingsData(listings))
            {
                ListingsDataGrid.ItemsSource = listings;
            }

            ListingCountText.Text = showCachedLabel
                ? $"{listings.Count} listings (cached)"
                : $"{listings.Count} listings";
            UpdateListingProgress(listings);
            return true;
        }

        private string SafeGetString(DbDataReader reader, int columnIndex)
        {
            if (reader.IsDBNull(columnIndex))
                return string.Empty;

            var value = reader.GetValue(columnIndex);
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private void UpdateListingProgress(IReadOnlyCollection<ListingViewModel> listings)
        {
            int total = listings?.Count ?? 0;
            int processed = 0;
            int successful = 0;
            int notSuccessful = 0;

            if (listings != null)
            {
                foreach (var listing in listings)
                {
                    if (!IsListingProcessed(listing))
                    {
                        continue;
                    }

                    processed++;
                    if (IsListingSuccessful(listing))
                    {
                        successful++;
                    }
                    else
                    {
                        notSuccessful++;
                    }
                }

                if (processed > total)
                {
                    processed = total;
                }
            }

            ListingProgressBar.Maximum = Math.Max(1, total);
            ListingProgressBar.Value = processed;
            ListingProgressText.Text = $"{processed}/{total}";
            SuccessfulCountText.Text = $"✅: {successful}";
            NotSuccessfulCountText.Text = $"⛔: {notSuccessful}";
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

        private static bool IsListingSuccessful(ListingViewModel listing)
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

            return status.Equals("Available", StringComparison.OrdinalIgnoreCase) ||
                   status.Equals("Not Available", StringComparison.OrdinalIgnoreCase);
        }

        private bool ShouldReplaceListingsData(IReadOnlyList<ListingViewModel> nextListings)
        {
            if (ListingsDataGrid?.ItemsSource is not IEnumerable<ListingViewModel> currentEnumerable)
            {
                return true;
            }

            var currentListings = currentEnumerable.ToList();

            if (currentListings.Count != nextListings.Count)
            {
                return true;
            }

            for (int i = 0; i < currentListings.Count; i++)
            {
                var current = currentListings[i];
                var next = nextListings[i];

                if (!string.Equals(current.ListingId, next.ListingId, StringComparison.Ordinal) ||
                    !string.Equals(current.CaseNumber, next.CaseNumber, StringComparison.Ordinal) ||
                    !string.Equals(current.Platform, next.Platform, StringComparison.Ordinal) ||
                    !string.Equals(current.Url, next.Url, StringComparison.Ordinal) ||
                    !string.Equals(current.UrlStatus, next.UrlStatus, StringComparison.Ordinal) ||
                    !string.Equals(current.CheckedDate, next.CheckedDate, StringComparison.Ordinal) ||
                    !string.Equals(current.Notes, next.Notes, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        //private bool CheckForChanges(List<RequestViewModel> newRequests)
        //{
        //    if (_allRequests.Count != newRequests.Count)
        //        return true;

        //    for (int i = 0; i < _allRequests.Count; i++)
        //    {
        //        var oldRequest = _allRequests[i];
        //        var newRequest = newRequests.FirstOrDefault(r => r.Id == oldRequest.Id);

        //        if (newRequest == null)
        //            return true;

        //        if (oldRequest.Status != newRequest.Status ||
        //            oldRequest.ListingsCount != newRequest.ListingsCount ||
        //            oldRequest.Priority != newRequest.Priority ||
        //            !string.Equals(oldRequest.Environment, newRequest.Environment, StringComparison.OrdinalIgnoreCase))
        //        {
        //            return true;
        //        }
        //    }

        //    foreach (var newRequest in newRequests)
        //    {
        //        if (!_allRequests.Any(r => r.Id == newRequest.Id))
        //            return true;
        //    }

        //    return false;
        //}

        //private void UpdateRequestsList(List<RequestViewModel> newRequests)
        //{
        //    foreach (var existingRequest in _allRequests.ToList())
        //    {
        //        var updatedRequest = newRequests.FirstOrDefault(r => r.Id == existingRequest.Id);
        //        if (updatedRequest != null)
        //        {
        //            existingRequest.Status = updatedRequest.Status;
        //            existingRequest.ListingsCount = updatedRequest.ListingsCount;
        //            existingRequest.StatusBrush = GetStatusBrush(existingRequest.Status);
        //            existingRequest.RequestInfoId = updatedRequest.RequestInfoId;
        //            existingRequest.Priority = updatedRequest.Priority;
        //            existingRequest.IsHighPriority = updatedRequest.IsHighPriority;
        //            existingRequest.Environment = updatedRequest.Environment;
        //            existingRequest.EnvironmentBrush = GetEnvironmentBrush(updatedRequest.Environment);
        //        }
        //    }

        //    var pendingRequests = 0;
        //    var newRequestIds = newRequests.Select(r => r.Id).Except(_allRequests.Select(r => r.Id));
        //    foreach (var newId in newRequestIds)
        //    {
        //        var request = newRequests.First(r => r.Id == newId);
        //        var requestVM = new RequestViewModel
        //        {
        //            Id = request.Id,
        //            RequestInfoId = request.RequestInfoId,
        //            User = request.User,
        //            FileName = request.FileName,
        //            Environment = NormalizeEnvironment(request.Environment),
        //            Status = request.Status,
        //            CreatedAt = request.CreatedAt,
        //            ListingsCount = request.ListingsCount,
        //            StatusBrush = GetStatusBrush(request.Status),
        //            Priority = request.Priority,
        //            IsHighPriority = request.Priority == 1,
        //            EnvironmentBrush = GetEnvironmentBrush(request.Environment)
        //        };

        //        if (request.Status == RequestStatus.PENDING)
        //        {
        //            pendingRequests++;
        //        }
        //        _allRequests.Add(requestVM);
        //    }

        //    _allRequests = _allRequests.OrderByDescending(r => r.Id).ToList();
        //    if (!string.IsNullOrWhiteSpace(SearchTextBox?.Text))
        //    {
        //        UpdateRequestCountText();
        //    }
        //    else
        //    {
        //        RequestCountText.Text = $"({pendingRequests} pending)";
        //    }
        //}

        //private bool IsFinalStatus(RequestStatus status)
        //{
        //    return status == RequestStatus.SUCCESS ||
        //           status == RequestStatus.FAILED ||
        //           status == RequestStatus.COMPLETED_WITH_ISSUES;
        //}

        //private void ShowRefreshNotification()
        //{
        //    // This method is intentionally left empty as it was only updating a variable that wasn't used
        //}

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
                    request.Status.ToString().ToLower().Contains(searchText) ||
                    request.Environment.ToLower().Contains(searchText);

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

        private static string NormalizeEnvironment(string? environment)
        {
            if (!string.IsNullOrWhiteSpace(environment) &&
                environment.Equals("Live", StringComparison.OrdinalIgnoreCase))
            {
                return "Live";
            }

            return "Stage";
        }

        private static SolidColorBrush GetEnvironmentBrush(string? environment)
        {
            var normalizedEnvironment = NormalizeEnvironment(environment);
            return normalizedEnvironment == "Live"
                ? new SolidColorBrush(Color.FromRgb(16, 185, 129))
                : new SolidColorBrush(Color.FromRgb(245, 158, 11));
        }

        private SolidColorBrush CreateBlinkingBrush(Color baseColor)
        {
            var brush = new SolidColorBrush(baseColor);

            // blinking effect
            var opacityAnimation = new DoubleAnimation
            {
                From = 0.4,
                To = 1.0,
                Duration = TimeSpan.FromSeconds(0.8),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
            };

            brush.BeginAnimation(SolidColorBrush.OpacityProperty, opacityAnimation);

            return brush;
        }

        private void RequestsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RequestsListBox.SelectedItem is RequestViewModel selectedRequest)
            {
                long? previousSelectedId = _selectedRequest?.Id;
                _selectedRequest = selectedRequest;

                bool showSkeleton = !_isRefreshing && !_suppressNextListingSkeleton;
                _suppressNextListingSkeleton = false;

                if (_isRefreshing && previousSelectedId.HasValue && previousSelectedId.Value == selectedRequest.Id)
                {
                    _suppressNextListingReload = false;
                    UpdateSelectedRequestDisplay(selectedRequest);
                    ExportButton.IsEnabled = IsExportAllowed(selectedRequest.Status);
                    return;
                }

                if (_suppressNextListingReload)
                {
                    _suppressNextListingReload = false;
                    if (ListingsDataGrid?.ItemsSource == null)
                    {
                        _ = LoadListingsForRequestAsync(selectedRequest.Id, showSkeleton: false);
                    }
                }
                else
                {
                    _ = LoadListingsForRequestAsync(selectedRequest.Id, showSkeleton: showSkeleton);
                }

                UpdateSelectedRequestDisplay(selectedRequest);
                ExportButton.IsEnabled = IsExportAllowed(selectedRequest.Status);
            }
            else
            {
                if (_isRestoringRequestSelection || _isRefreshing)
                {
                    return;
                }

                _suppressNextListingReload = false;
                _suppressNextListingSkeleton = false;
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
            SelectedRequestInfo.Text = $"RID: {request.RequestInfoId} • User: {request.User} • Created: {request.CreatedAt:yyyy MMM d, h:mm tt}";
            SelectedRequestStatus.Text = $"Status: {request.Status}";
            PriorityToggleButton.IsChecked = request.IsHighPriority;
            PriorityToggleButton.IsEnabled = true;
            PriorityToggleButton.ToolTip = request.IsHighPriority ? "Click to remove high priority" : "Click to set high priority";

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
                ShowSuccessMessage(setHighPriority ? "Request marked as high priority." : "Request priority cleared.", "Priority Updated");
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Error updating priority: {ex.Message}", "Priority Update Failed");
            }
        }

        private void ClearListings()
        {
            SetListingsLoadingState(false);
            ListingsDataGrid.ItemsSource = null;
            ListingCountText.Text = "0 listings";
            UpdateListingProgress(null);
        }

        private void SetListingsLoadingState(bool isLoading)
        {
            if (_isListingsLoading == isLoading)
            {
                return;
            }

            _isListingsLoading = isLoading;

            if (ListingsSkeletonOverlay != null)
            {
                ListingsSkeletonOverlay.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            }

            if (ListingsDataGrid != null)
            {
                ListingsDataGrid.Visibility = isLoading ? Visibility.Hidden : Visibility.Visible;
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _currentPage = 1;
            _searchDebounceTimer.Stop();
            _searchDebounceTimer.Start();
        }

        private async void SearchDebounceTimer_Tick(object sender, EventArgs e)
        {
            _searchDebounceTimer.Stop();
            if (!_isConnectionAvailable)
            {
                TryLoadRequestsFromCache(preserveSelection: false);
                return;
            }

            await LoadRequestsPageAsync(preserveSelection: false, showErrors: false);
        }

        private async void PrevPageButton_Click(object sender, RoutedEventArgs e)
        {
            await NavigateToPageAsync(_currentPage - 1);
        }

        private async void NextPageButton_Click(object sender, RoutedEventArgs e)
        {
            await NavigateToPageAsync(_currentPage + 1);
        }

        private async void FirstPageButton_Click(object sender, RoutedEventArgs e)
        {
            await NavigateToPageAsync(1);
        }

        private async void LastPageButton_Click(object sender, RoutedEventArgs e)
        {
            await NavigateToPageAsync(_totalPages);
        }

        private async void GoToPageButton_Click(object sender, RoutedEventArgs e)
        {
            await GoToPageFromInputAsync();
        }

        private async void PageSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || PageSizeComboBox?.SelectedItem is not ComboBoxItem selectedItem)
            {
                return;
            }

            if (!int.TryParse(selectedItem.Tag?.ToString(), out int selectedPageSize) || selectedPageSize <= 0)
            {
                return;
            }

            if (selectedPageSize == _pageSize)
            {
                return;
            }

            _pageSize = selectedPageSize;
            _currentPage = 1;
            await LoadRequestsPageAsync(preserveSelection: false, showErrors: false, showSkeleton: false);
        }

        private void PageJumpTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = e.Text.Any(c => !char.IsDigit(c));
        }

        private async void PageJumpTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
            {
                return;
            }

            e.Handled = true;
            await GoToPageFromInputAsync();
        }

        private async Task GoToPageFromInputAsync()
        {
            if (PageJumpTextBox == null)
            {
                return;
            }

            if (!int.TryParse(PageJumpTextBox.Text, out int targetPage))
            {
                PageJumpTextBox.Text = _currentPage.ToString(CultureInfo.InvariantCulture);
                return;
            }

            targetPage = Math.Max(1, Math.Min(targetPage, _totalPages));
            await NavigateToPageAsync(targetPage);
        }

        private void UpdateRequestCountText()
        {
            _currentSearchText = SearchTextBox?.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(_currentSearchText))
            {
                RequestCountText.Text = $"({_pendingRequests} pending)";
                return;
            }

            RequestCountText.Text = $"({_totalRequests} requests)";
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

            _rescanOnlyErrors = true;
            SelectRescanOption(true);

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
                    RescanId = _rescanOnlyErrors ? 1 : 0,
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

            int listingsCount = 0;
            if (ListingsDataGrid.ItemsSource is System.Collections.IEnumerable items)
            {
                listingsCount = items.Cast<object>().Count();
            }

            ExportFileName.Text = Path.GetFileNameWithoutExtension(_selectedRequest.FileName);
            ExportListingsCount.Text = listingsCount.ToString();
            ExportStatus.Text = _selectedRequest.Status.ToString();

            string baseFileName = Path.GetFileNameWithoutExtension(_selectedRequest.FileName);
            string formattedDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string defaultFileName = $"{baseFileName}-Result-{formattedDate}.xlsx";
            string defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            _selectedExportPath = Path.Combine(defaultFolder, defaultFileName);
            ExportFilePath.Text = _selectedExportPath;

            ConfirmExportButton.IsEnabled = true;

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
                string directory = Path.GetDirectoryName(_selectedExportPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                ExportToExcel(_selectedRequest.Id, _selectedExportPath);

                ShowSuccessMessage($"Listings exported successfully.", "Export Complete");
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Error exporting to Excel: {ex.Message}", "Export Error");
            }
        }

        private void ExportToExcel(long requestId, string filePath)
        {
            try
            {
                using var context = new ProductCheckerDbContext();

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

        private static async Task<T> WithTimeoutAsync<T>(Func<CancellationToken, Task<T>> operation, int timeoutSeconds)
        {
            using var cts = new CancellationTokenSource();
            var task = operation(cts.Token);
            var completedTask = await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(timeoutSeconds))).ConfigureAwait(true);

            if (completedTask != task)
            {
                cts.Cancel();
                throw new TimeoutException($"The operation timed out after {timeoutSeconds} seconds.");
            }

            return await task.ConfigureAwait(true);
        }

        private static string GetCacheFilePath()
        {
            var baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ProductCheckerV2",
                "cache");

            return Path.Combine(baseDir, "view-requests-cache.json");
        }

        private void LoadCacheFromDisk()
        {
            try
            {
                if (!File.Exists(_cacheFilePath))
                {
                    _cacheData = new ViewRequestsCacheData();
                    return;
                }

                var json = File.ReadAllText(_cacheFilePath);
                var data = JsonSerializer.Deserialize<ViewRequestsCacheData>(json);
                _cacheData = data ?? new ViewRequestsCacheData();
            }
            catch
            {
                _cacheData = new ViewRequestsCacheData();
            }
        }

        private void SaveCacheToDisk()
        {
            try
            {
                lock (_cacheLock)
                {
                    var dir = Path.GetDirectoryName(_cacheFilePath);
                    if (!string.IsNullOrWhiteSpace(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    _cacheData.UpdatedAt = DateTime.UtcNow;
                    var json = JsonSerializer.Serialize(_cacheData, new JsonSerializerOptions
                    {
                        WriteIndented = false
                    });
                    File.WriteAllText(_cacheFilePath, json);
                }
            }
            catch
            {
                // Ignore cache write failures.
            }
        }

        private bool TryLoadRequestsFromCache(bool preserveSelection)
        {
            List<CachedRequest> cached;
            Dictionary<long, List<CachedListing>> cachedListingsByRequestId;
            lock (_cacheLock)
            {
                cached = _cacheData.Requests.ToList();
                cachedListingsByRequestId = _cacheData.ListingsByRequestId
                    .ToDictionary(
                        kvp => kvp.Key,
                        kvp => kvp.Value.ToList());
            }

            if (cached.Count == 0)
            {
                return false;
            }

            var search = SearchTextBox?.Text?.Trim() ?? string.Empty;
            var filtered = cached
                .Select(r =>
                {
                    var requestMatches = CachedRequestMatchesSearch(r, search);
                    var listingMatchCount = GetCachedListingMatchCount(r.Id, search, cachedListingsByRequestId);
                    return new
                    {
                        Request = r,
                        ListingMatchCount = listingMatchCount,
                        Matches = requestMatches || listingMatchCount > 0
                    };
                })
                .Where(x => x.Matches)
                .OrderByDescending(x => x.Request.Id)
                .ToList();

            _totalRequests = filtered.Count;
            _pendingRequests = filtered.Count(r => string.Equals(r.Request.Status, RequestStatus.PENDING.ToString(), StringComparison.OrdinalIgnoreCase));
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)Math.Max(1, _totalRequests) / _pageSize));
            _currentPage = Math.Max(1, Math.Min(_currentPage, _totalPages));

            var page = filtered
                .Skip((_currentPage - 1) * _pageSize)
                .Take(_pageSize)
                .Select(x => ToRequestViewModel(x.Request, x.ListingMatchCount))
                .ToList();

            _allRequests = page;
            _requestsView = CollectionViewSource.GetDefaultView(_allRequests);
            _requestsView.Filter = RequestFilter;
            RequestsListBox.ItemsSource = _requestsView;
            _requestsView.Refresh();

            UpdateRequestCountText();
            UpdatePaginationControls();
            ApplySelectionAfterLoad(preserveSelection);

            return true;
        }

        private bool TryLoadListingsFromCache(long requestId, out List<ListingViewModel> listings)
        {
            lock (_cacheLock)
            {
                if (_cacheData.ListingsByRequestId.TryGetValue(requestId, out var cached))
                {
                    listings = cached.Select(c => new ListingViewModel
                    {
                        ListingId = c.ListingId,
                        CaseNumber = c.CaseNumber,
                        Platform = c.Platform,
                        Url = c.Url,
                        UrlStatus = c.UrlStatus,
                        CheckedDate = c.CheckedDate,
                        Notes = c.Notes
                    }).ToList();

                    return true;
                }
            }

            listings = new List<ListingViewModel>();
            return false;
        }

        private static bool CachedRequestMatchesSearch(CachedRequest request, string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return true;
            }

            var search = searchText.ToLowerInvariant();
            return request.Id.ToString().Contains(search, StringComparison.OrdinalIgnoreCase) ||
                   request.RequestInfoId.ToString().Contains(search, StringComparison.OrdinalIgnoreCase) ||
                   (request.User ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                   (request.FileName ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                   (request.Environment ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                   (request.Status ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase);
        }

        private static int GetCachedListingMatchCount(
            long requestId,
            string searchText,
            IReadOnlyDictionary<long, List<CachedListing>> cachedListingsByRequestId)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return 0;
            }

            if (!cachedListingsByRequestId.TryGetValue(requestId, out var listings) || listings == null || listings.Count == 0)
            {
                return 0;
            }

            var search = searchText.Trim();
            return listings.Count(l =>
                (l.ListingId ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                (l.CaseNumber ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                (l.Platform ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                (l.Url ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                (l.UrlStatus ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                (l.CheckedDate ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase) ||
                (l.Notes ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase));
        }

        private RequestViewModel ToRequestViewModel(CachedRequest cached, int listingMatchCount)
        {
            var status = Enum.TryParse<RequestStatus>(cached.Status, true, out var parsedStatus)
                ? parsedStatus
                : RequestStatus.PENDING;

            return new RequestViewModel
            {
                Id = cached.Id,
                RequestInfoId = cached.RequestInfoId,
                User = cached.User ?? "Unknown",
                FileName = cached.FileName ?? "Unknown",
                Environment = NormalizeEnvironment(cached.Environment),
                Status = status,
                CreatedAt = cached.CreatedAt,
                ListingsCount = cached.ListingsCount,
                Priority = cached.Priority,
                IsHighPriority = cached.Priority == 1,
                MatchCount = listingMatchCount,
                HasListingMatch = listingMatchCount > 0,
                StatusBrush = GetStatusBrush(status),
                EnvironmentBrush = GetEnvironmentBrush(cached.Environment)
            };
        }

        private sealed class ViewRequestsCacheData
        {
            public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
            public List<CachedRequest> Requests { get; set; } = new();
            public Dictionary<long, List<CachedListing>> ListingsByRequestId { get; set; } = new();
        }

        private sealed class CachedRequest
        {
            public long Id { get; set; }
            public long RequestInfoId { get; set; }
            public string User { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
            public string Environment { get; set; } = "Stage";
            public string Status { get; set; } = RequestStatus.PENDING.ToString();
            public DateTime CreatedAt { get; set; }
            public int ListingsCount { get; set; }
            public int Priority { get; set; }
        }

        private sealed class CachedListing
        {
            public string ListingId { get; set; } = string.Empty;
            public string CaseNumber { get; set; } = string.Empty;
            public string Platform { get; set; } = string.Empty;
            public string Url { get; set; } = string.Empty;
            public string UrlStatus { get; set; } = string.Empty;
            public string CheckedDate { get; set; } = string.Empty;
            public string Notes { get; set; } = string.Empty;
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
    }
}



