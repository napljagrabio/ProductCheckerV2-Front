using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using ProductCheckerV2.Common;
using ProductCheckerV2.Artemis;
using ProductCheckerV2.Database;
using ProductCheckerV2.Database.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Animation;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using System.Windows.Data;

namespace ProductCheckerV2
{
    public partial class MainWindow : Window
    {
        private List<UploadedProductData> _uploadedData = new List<UploadedProductData>();
        private BackgroundWorker _processingWorker;
        private string _applicationName = "";
        private string _currentFilePath = string.Empty;
        private int _currentRequestId = 0;
        private int _currentRecordsProcessed = 0;
        private static List<CrawlerPlatform> _platformsCache;
        private bool _isFileValidationActive = false;
        private bool _isDragOver = false;
        private bool _useFilterMode = true;
        private List<FilterOption> _campaignOptions = new List<FilterOption>();
        private List<FilterOption> _caseOptions = new List<FilterOption>();
        private List<FilterOption> _platformOptions = new List<FilterOption>();
        private List<FilterOption> _qflagOptions = new List<FilterOption>();
        private ICollectionView _campaignsView;
        private ICollectionView _casesView;
        private ICollectionView _platformsView;
        private ICollectionView _qflagsView;
        private ObservableCollection<SelectedFilterItem> _selectedFilters = new ObservableCollection<SelectedFilterItem>();
        private string _customFileName = string.Empty;
        private bool _isEnvironmentSelectorInitialized = false;
        private string _pendingEnvironment = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
            InitializeApp();
            SetupBackgroundWorker();
            SetupDragDrop();
        }

                private void InitializeApp()
        {
            try
            {
                _applicationName = ConfigurationManager.ApplicationName;
                this.Icon = new BitmapImage(new Uri("Assets/Logo.ico".AbsPath()));
                LogoImage.Source = new BitmapImage(new Uri("Assets/Logo.ico".AbsPath()));

                InitializeEnvironmentSelector();
                UpdateEnvironmentIndicator();
                InitializeDatabase();
                LoadFilterOptions();
                SetInputMode(FilterModeRadio?.IsChecked == true);
                UpdateSelectionSummary();
                UpdateStatusBar("Ready");
                UpdateStatsDisplay();
                UpdateDataGridVisibility(false);
                ShowUploadPage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Application initialization failed:\n\n{ex.Message}",
                    "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InitializeEnvironmentSelector()
        {
            var configuredEnvironment = ConfigurationManager.GetEnvironment();
            ConfigurationManager.SetEnvironment(configuredEnvironment);
            SetEnvironmentComboBoxSelection(configuredEnvironment);
            _pendingEnvironment = configuredEnvironment;
            _isEnvironmentSelectorInitialized = true;
        }

        private void UpdateEnvironmentIndicator()
        {
            var environment = ConfigurationManager.GetEnvironment();
            var isStage = environment.Equals("Stage", StringComparison.OrdinalIgnoreCase);

            if (StageIndicatorPanel != null)
            {
                StageIndicatorPanel.Visibility = isStage ? Visibility.Visible : Visibility.Collapsed;
            }

            if (StageIndicatorText != null)
            {
                StageIndicatorText.Text = environment;
            }
        }

        private void SetEnvironmentComboBoxSelection(string environment)
        {
            var targetEnvironment = environment.Equals("Live", StringComparison.OrdinalIgnoreCase)
                ? "Live"
                : "Stage";

            var previousGuardState = _isEnvironmentSelectorInitialized;
            _isEnvironmentSelectorInitialized = false;

            foreach (var item in SettingsEnvironmentComboBox.Items.OfType<ComboBoxItem>())
            {
                var itemValue = item.Content?.ToString();
                item.IsSelected = string.Equals(itemValue, targetEnvironment, StringComparison.OrdinalIgnoreCase);
            }

            _isEnvironmentSelectorInitialized = previousGuardState;
        }

        private void ApplyEnvironmentChange(string selectedEnvironment)
        {
            var previousEnvironment = ConfigurationManager.GetEnvironment();

            try
            {
                ConfigurationManager.SetEnvironment(selectedEnvironment);

                _platformsCache = null;
                _selectedFilters.Clear();

                InitializeDatabase();
                LoadFilterOptions();
                UpdateSelectionSummary();
                UpdateStatsDisplay();

                UpdateStatusBar($"Environment: {selectedEnvironment}");
                UpdateEnvironmentIndicator();
            }
            catch (Exception ex)
            {
                ConfigurationManager.SetEnvironment(previousEnvironment);
                SetEnvironmentComboBoxSelection(previousEnvironment);
                UpdateEnvironmentIndicator();

                MessageBox.Show($"Failed to switch environment to '{selectedEnvironment}'.\n\n{ex.Message}",
                    "Environment Switch Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EnvironmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isEnvironmentSelectorInitialized)
            {
                return;
            }

            var selectedEnvironment = (SettingsEnvironmentComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (string.IsNullOrWhiteSpace(selectedEnvironment))
            {
                return;
            }

            _pendingEnvironment = selectedEnvironment;
        }

        private void SettingsToggleButton_Click(object sender, RoutedEventArgs e)
        {
            var isOpen = SettingsToggleButton?.IsChecked == true;

            if (isOpen)
            {
                var currentEnvironment = ConfigurationManager.GetEnvironment();
                SetEnvironmentComboBoxSelection(currentEnvironment);
                _pendingEnvironment = currentEnvironment;

                if (EnvironmentPasswordBox != null)
                {
                    EnvironmentPasswordBox.Password = string.Empty;
                    EnvironmentPasswordBox.Focus();
                }
            }

            if (EnvironmentSettingsPopup != null)
            {
                EnvironmentSettingsPopup.IsOpen = isOpen;
            }
        }

        private void EnvironmentSettingsPopup_Closed(object sender, EventArgs e)
        {
            if (SettingsToggleButton != null)
            {
                SettingsToggleButton.IsChecked = false;
            }
        }

        private void ApplyEnvironmentSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isEnvironmentSelectorInitialized)
            {
                return;
            }

            var selectedEnvironment = _pendingEnvironment;
            if (string.IsNullOrWhiteSpace(selectedEnvironment))
            {
                selectedEnvironment = (SettingsEnvironmentComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(selectedEnvironment))
            {
                MessageBox.Show("Please select an environment first.",
                    "Environment Required", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var expectedPassword = ConfigurationManager.GetEnvironmentSwitchPassword();
            var enteredPassword = EnvironmentPasswordBox?.Password ?? string.Empty;

            if (!string.Equals(enteredPassword, expectedPassword, StringComparison.Ordinal))
            {
                MessageBox.Show("Invalid password. Environment was not changed.",
                    "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);

                if (EnvironmentPasswordBox != null)
                {
                    EnvironmentPasswordBox.Password = string.Empty;
                    EnvironmentPasswordBox.Focus();
                }
                return;
            }

            if (!string.Equals(selectedEnvironment, ConfigurationManager.GetEnvironment(), StringComparison.OrdinalIgnoreCase))
            {
                ApplyEnvironmentChange(selectedEnvironment);
            }

            if (EnvironmentPasswordBox != null)
            {
                EnvironmentPasswordBox.Password = string.Empty;
            }
            if (EnvironmentSettingsPopup != null)
            {
                EnvironmentSettingsPopup.IsOpen = false;
            }
        }

        private void InitializeDatabase()
        {
            try
            {
                using var context = new ProductCheckerDbContext();
                context.Database.EnsureCreated();

                if (context.Database.CanConnect())
                {
                    Console.WriteLine("Database connected successfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database initialization failed:\n\n{ex.Message}",
                    "Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadFilterOptions()
        {
            try
            {
                using var context = new ArtemisDbContext();

                _campaignOptions = context.Campaigns
                    .Where(c => c.DeletedAt == null && c.Status == 1)
                    .OrderBy(c => c.Name)
                    .Select(c => new FilterOption
                    {
                        Id = c.Id,
                        Display = string.IsNullOrWhiteSpace(c.Name) ? $"Campaign #{c.Id}" : c.Name
                    })
                    .ToList();

                _caseOptions = context.Cases
                    .Where(c => c.DeletedAt == null)
                    .OrderBy(c => c.CaseNumber)
                    .Select(c => new FilterOption
                    {
                        Id = c.Id,
                        Display = string.IsNullOrWhiteSpace(c.CaseNumber) ? $"Case #{c.Id}" : c.CaseNumber
                    })
                    .ToList();

                _platformOptions = context.Platforms
                    .Where(p => p.DeletedAt == null && p.Status == 1)
                    .OrderBy(p => p.Name)
                    .Select(p => new FilterOption
                    {
                        Id = p.Id,
                        Display = string.IsNullOrWhiteSpace(p.Name) ? $"Platform #{p.Id}" : p.Name
                    })
                    .ToList();

                _qflagOptions = context.Qflag
                    .Where(q => q.DeletedAt == null && q.Status == 1)
                    .OrderBy(q => q.Label)
                    .Select(q => new FilterOption
                    {
                        Id = q.Id,
                        Display = string.IsNullOrWhiteSpace(q.Label) ? $"QFlag #{q.Id}" : q.Label
                    })
                    .ToList();

                CampaignListBox.ItemsSource = _campaignOptions;
                CaseListBox.ItemsSource = _caseOptions;
                PlatformListBox.ItemsSource = _platformOptions;
                QflagListBox.ItemsSource = _qflagOptions;
                SelectedFiltersListBox.ItemsSource = _selectedFilters;

                _campaignsView = CollectionViewSource.GetDefaultView(_campaignOptions);
                _campaignsView.Filter = item => FilterOptionMatches(item, CampaignSearchTextBox.Text);
                ConfigureFilterView(_campaignsView);

                _casesView = CollectionViewSource.GetDefaultView(_caseOptions);
                _casesView.Filter = item => FilterOptionMatches(item, CaseSearchTextBox.Text);
                ConfigureFilterView(_casesView);

                _platformsView = CollectionViewSource.GetDefaultView(_platformOptions);
                _platformsView.Filter = item => FilterOptionMatches(item, PlatformSearchTextBox.Text);
                ConfigureFilterView(_platformsView);

                _qflagsView = CollectionViewSource.GetDefaultView(_qflagOptions);
                _qflagsView.Filter = item => FilterOptionMatches(item, QflagSearchTextBox.Text);
                ConfigureFilterView(_qflagsView);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading filters:\n\n{ex.Message}",
                    "Filter Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static bool FilterOptionMatches(object item, string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return true;
            }

            if (item is FilterOption option)
            {
                return option.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                       option.Id.ToString().Contains(searchText, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private static void ConfigureFilterView(ICollectionView view)
        {
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(nameof(FilterOption.IsSelected), ListSortDirection.Descending));
            view.SortDescriptions.Add(new SortDescription(nameof(FilterOption.Display), ListSortDirection.Ascending));

            if (view is ICollectionViewLiveShaping liveView)
            {
                liveView.LiveSortingProperties.Clear();
                liveView.LiveSortingProperties.Add(nameof(FilterOption.IsSelected));
                liveView.LiveSortingProperties.Add(nameof(FilterOption.Display));
                liveView.IsLiveSorting = true;
            }
        }

        private void CampaignSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _campaignsView?.Refresh();
        }

        private void CaseSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _casesView?.Refresh();
        }

        private void PlatformSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _platformsView?.Refresh();
        }

        private void QflagSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _qflagsView?.Refresh();
        }

        private void FilterOptionCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            UpdateSelectionSummary();
            RefreshFilterViews();
        }

        private void ModeRadio_Checked(object sender, RoutedEventArgs e)
        {
            SetInputMode(FilterModeRadio.IsChecked == true);
        }

        private void SetInputMode(bool useFilterMode)
        {
            _useFilterMode = useFilterMode;

            if (FiltersPanel != null)
            {
                FiltersPanel.Visibility = useFilterMode ? Visibility.Visible : Visibility.Collapsed;
            }
            if (UploadPanel != null)
            {
                UploadPanel.Visibility = useFilterMode ? Visibility.Collapsed : Visibility.Visible;
            }

            UpdateStatusBar(useFilterMode
                ? "Ready - Select filters and preview listings"
                : "Ready - Upload Excel file");
        }

        private void RemoveSelectedFilter_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is SelectedFilterItem item)
            {
                item.Option.IsSelected = false;
                UpdateSelectionSummary();
                RefreshFilterViews();
            }
        }

        private void RefreshFilterViews()
        {
            _campaignsView?.Refresh();
            _casesView?.Refresh();
            _platformsView?.Refresh();
            _qflagsView?.Refresh();
        }

        private async Task LoadPreviewListingsAsync()
        {
            if (!_useFilterMode)
            {
                return;
            }

            UpdateSelectionSummary();

            var selectedCampaignIds = GetSelectedIds(_campaignOptions);
            var selectedCaseIds = GetSelectedIds(_caseOptions);
            var selectedPlatformIds = GetSelectedIds(_platformOptions);
            var selectedQflagIds = GetSelectedIds(_qflagOptions);

            if (selectedCampaignIds.Count == 0 &&
                selectedCaseIds.Count == 0 &&
                selectedPlatformIds.Count == 0 &&
                selectedQflagIds.Count == 0)
            {
                MessageBox.Show("Please select at least one campaign, case, platform, or status.",
                    "Missing Filters", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            BrowseButton.IsEnabled = false;
            StartButton.IsEnabled = false;
            UpdateStatusBar("Loading preview...");

            try
            {
                var listings = await Task.Run(() =>
                    QueryListings(selectedCampaignIds, selectedCaseIds, selectedPlatformIds, selectedQflagIds));

                _uploadedData = listings;
                DataGridPreview.ItemsSource = _uploadedData;
                UpdateDataGridVisibility(_uploadedData.Count > 0);
                RecordCountText.Text = $"{_uploadedData.Count} records loaded";
                StartButton.IsEnabled = _uploadedData.Count > 0;

                UpdateStatusBar(_uploadedData.Count > 0
                    ? "Preview loaded"
                    : "No listings found for the selected filters");
                if (_uploadedData.Count == 0)
                {
                    MessageBox.Show("No listings found for the selected filters.",
                        "No Results", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading preview:\n\n{ex.Message}",
                    "Preview Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BrowseButton.IsEnabled = true;
            }
        }

        private static List<UploadedProductData> QueryListings(
            List<int> campaignIds,
            List<int> caseIds,
            List<int> platformIds,
            List<int> qflagIds)
        {
            using var context = new ArtemisDbContext();

            var query = from listing in context.Listings
                        join c in context.Cases on listing.CaseId equals c.Id
                        join p in context.Platforms on listing.PlatformId equals p.Id
                        join q in context.Qflag on listing.QfalgId equals q.Id
                        where listing.DeletedAt == null &&
                              c.DeletedAt == null &&
                              p.DeletedAt == null &&
                              q.DeletedAt == null
                        select new
                        {
                            listing.Id,
                            listing.CampaignId,
                            listing.CaseId,
                            listing.PlatformId,
                            listing.QfalgId,
                            CaseNumber = c.CaseNumber,
                            PlatformName = p.Name,
                            Url = listing.Url
                        };

            if (campaignIds.Count > 0)
            {
                query = query.Where(item => campaignIds.Contains(item.CampaignId));
            }

            if (caseIds.Count > 0)
            {
                query = query.Where(item => caseIds.Contains(item.CaseId));
            }

            if (platformIds.Count > 0)
            {
                query = query.Where(item => platformIds.Contains(item.PlatformId));
            }

            if (qflagIds.Count > 0)
            {
                query = query.Where(item => qflagIds.Contains(item.QfalgId));
            }

            return query.ToList()
                .Select(item => new UploadedProductData
                {
                    ListingId = item.Id.ToString(),
                    CaseNumber = item.CaseNumber,
                    ProductUrl = item.Url ?? string.Empty,
                    Platform = GetPlatformFromUrl(item.Url ?? string.Empty)
                })
                .ToList();
        }

        private static List<int> GetSelectedIds(List<FilterOption> options)
        {
            return options
                .Where(option => option.IsSelected)
                .Select(option => option.Id)
                .ToList();
        }

        private void UpdateSelectionSummary()
        {
            int campaignCount = _campaignOptions.Count(o => o.IsSelected);
            int caseCount = _caseOptions.Count(o => o.IsSelected);
            int platformCount = _platformOptions.Count(o => o.IsSelected);
            int qflagCount = _qflagOptions.Count(o => o.IsSelected);

            if (SelectedCountsText != null)
            {
                SelectedCountsText.Text = $"Campaigns: {campaignCount}  |  Cases: {caseCount}  |  Platforms: {platformCount}  |  Status: {qflagCount}";
            }

            UpdateSelectedFiltersList();
        }

        private void UpdateSelectedFiltersList()
        {
            _selectedFilters.Clear();
            AddSelectedFilters("Campaign", _campaignOptions);
            AddSelectedFilters("Case", _caseOptions);
            AddSelectedFilters("Platform", _platformOptions);
            AddSelectedFilters("Status", _qflagOptions);
        }

        private void AddSelectedFilters(string category, List<FilterOption> options)
        {
            foreach (var option in options.Where(o => o.IsSelected))
            {
                _selectedFilters.Add(new SelectedFilterItem
                {
                    Category = category,
                    Option = option,
                    Display = $"{category}: {option.Display}"
                });
            }
        }

        private void ClearSelections()
        {
            ClearSelectionList(_campaignOptions);
            ClearSelectionList(_caseOptions);
            ClearSelectionList(_platformOptions);
            ClearSelectionList(_qflagOptions);

            CampaignSearchTextBox.Text = string.Empty;
            CaseSearchTextBox.Text = string.Empty;
            PlatformSearchTextBox.Text = string.Empty;
            QflagSearchTextBox.Text = string.Empty;

            RefreshFilterViews();
        }

        private static void ClearSelectionList(List<FilterOption> options)
        {
            foreach (var option in options)
            {
                option.IsSelected = false;
            }
        }

        private void SetupBackgroundWorker()
        {
            _processingWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = false
            };

            _processingWorker.DoWork += ProcessingWorker_DoWork;
            _processingWorker.ProgressChanged += ProcessingWorker_ProgressChanged;
            _processingWorker.RunWorkerCompleted += ProcessingWorker_RunWorkerCompleted;
        }

        private void SetupDragDrop()
        {
            this.AllowDrop = true;
            this.PreviewDragOver += MainWindow_PreviewDragOver;
            this.PreviewDrop += MainWindow_PreviewDrop;

            if (UploadArea != null)
            {
                UploadArea.PreviewDragOver += UploadArea_PreviewDragOver;
                UploadArea.PreviewDragLeave += UploadArea_PreviewDragLeave;
                UploadArea.PreviewDrop += UploadArea_PreviewDrop;
            }
        }

        private void MainWindow_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private async void MainWindow_PreviewDrop(object sender, DragEventArgs e)
        {
            if (!_useFilterMode && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    await ValidateAndLoadFileAsync(files[0]);
                }
            }
        }

        private void UploadArea_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (!_isDragOver && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                _isDragOver = true;
                UploadArea.Background = new SolidColorBrush(
                    Color.FromArgb(30, 67, 97, 238));
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
        }

        private void UploadArea_PreviewDragLeave(object sender, DragEventArgs e)
        {
            _isDragOver = false;
            UploadArea.Background = Brushes.Transparent;
        }

        private async void UploadArea_PreviewDrop(object sender, DragEventArgs e)
        {
            _isDragOver = false;
            UploadArea.Background = Brushes.Transparent;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    await ValidateAndLoadFileAsync(files[0]);
                }
            }
            e.Handled = true;
        }

        private async void PreviewListingsButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadPreviewListingsAsync();
        }
        private async void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // Ensure the browse button opens the file picker.
            UploadBrowseButton_Click(sender, e);
        }

        private async void UploadBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|All files (*.*)|*.*",
                Title = "Select Excel File",
                DefaultExt = ".xlsx",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog() == true)
            {
                await ValidateAndLoadFileAsync(dialog.FileName);
            }
        }

        

        private async Task ValidateAndLoadFileAsync(string filePath)
        {
            try
            {
                // Validate file size (max 10MB)
                var fileInfo = new FileInfo(filePath);
                long maxSize = 10 * 1024 * 1024;

                if (fileInfo.Length > maxSize)
                {
                    MessageBox.Show($"File size exceeds the maximum allowed size of 10MB.",
                        "File Too Large", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Validate extension
                var validExtensions = new[] { ".xlsx", ".xls", ".xlsm" };
                if (!validExtensions.Contains(fileInfo.Extension.ToLower()))
                {
                    MessageBox.Show("Please select a valid Excel file (.xlsx, .xls, .xlsm).",
                        "Invalid File Type", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                await LoadExcelFileAsync(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading file:\n\n{ex.Message}",
                    "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task LoadExcelFileAsync(string filePath)
        {
            ShowValidationOverlay("Validating Excel file...", "Reading rows and checking file structure...");
            try
            {
                _currentFilePath = filePath;
                if (UploadFilePathText != null)
                {
                    UploadFilePathText.Text = Path.GetFileName(filePath);
                }
                UpdateStatusBar($"Loading file: {Path.GetFileName(filePath)}...");

                // Read Excel data
                await Task.Yield();
                var data = await Task.Run(() => ReadExcelData(filePath));
                _uploadedData = data;

                if (_uploadedData.Count == 0)
                {
                    HideValidationOverlay();
                    MessageBox.Show("No valid data found in the Excel file.\n\n" +
                                  "Please ensure the file contains columns:\n" +
                                  "A: Listing ID, B: Case Number, C: URL",
                        "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearAllData();
                    return;
                }

                ProgressText.Text = "Validating listings...";
                ProgressSubText.Text = "Checking duplicates and verifying IDs...";
                var validationErrors = await ValidateListingsAsync(_uploadedData);
                if (validationErrors.Count > 0)
                {
                    HideValidationOverlay();
                    MessageBox.Show("Validation failed:\n\n" + string.Join("\n", validationErrors),
                        "Validation Errors", MessageBoxButton.OK, MessageBoxImage.Warning);

                    ClearAllData();
                    return;
                }

                DataGridPreview.ItemsSource = _uploadedData;
                UpdateDataGridVisibility(true);
                UpdateStatsDisplay();

                RecordCountText.Text = $"{_uploadedData.Count} records loaded";

                var platforms = _uploadedData.Select(d => d.Platform).Distinct().ToList();
                var infoItems = new List<FileInfoItem>
                {
                    new FileInfoItem {
                        Text = $"Total Records: {_uploadedData.Count}",
                        Color = new SolidColorBrush(Colors.Green)
                    },
                    new FileInfoItem {
                        Text = $"Platforms: {string.Join(", ", platforms)}",
                        Color = new SolidColorBrush(Colors.Blue)
                    },
                    new FileInfoItem {
                        Text = $"File Size: {FormatFileSize(new FileInfo(filePath).Length)}",
                        Color = new SolidColorBrush(Colors.Orange)
                    }
                };
                if (UploadFileInfoItems != null)
                {
                    UploadFileInfoItems.ItemsSource = infoItems;
                }

                StartButton.IsEnabled = true;
                UpdateStatusBar($"Ready - {_uploadedData.Count} records loaded");
            }
            catch (Exception ex)
            {
                HideValidationOverlay();
                MessageBox.Show($"Error loading Excel file:\n\n{ex.Message}",
                    "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearAllData();
            }
            finally
            {
                HideValidationOverlay();
                BrowseButton.IsEnabled = true;
                ClearButton.IsEnabled = true;
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double len = bytes;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        private List<UploadedProductData> ReadExcelData(string filePath)
        {
            var data = new List<UploadedProductData>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();
            var rows = worksheet.RowsUsed();

            if (!rows.Any()) return data;

            // Check if first row is header
            bool firstRowIsHeader = false;
            var firstRow = rows.First();

            var cell1 = firstRow.Cell(1).GetString().ToLower();
            var cell2 = firstRow.Cell(2).GetString().ToLower();
            var cell3 = firstRow.Cell(3).GetString().ToLower();

            if (cell1.Contains("listing") || cell1.Contains("id") ||
                cell2.Contains("case") || cell2.Contains("number") ||
                cell3.Contains("url") || cell3.Contains("link"))
            {
                firstRowIsHeader = true;
            }

            foreach (var row in rows)
            {
                if (firstRowIsHeader && row == rows.First())
                    continue;

                try
                {
                    // Use GetFormattedString() which handles all data types
                    var listingId = row.Cell(1).GetFormattedString().Trim();
                    var caseNumber = row.Cell(2).GetFormattedString().Trim();
                    var url = row.Cell(3).GetFormattedString().Trim();

                    // Skip if all empty
                    if (string.IsNullOrWhiteSpace(listingId) &&
                        string.IsNullOrWhiteSpace(caseNumber) &&
                        string.IsNullOrWhiteSpace(url))
                        continue;

                    data.Add(new UploadedProductData
                    {
                        ListingId = listingId,
                        CaseNumber = string.IsNullOrWhiteSpace(caseNumber) ? null : caseNumber,
                        ProductUrl = url,
                        Platform = GetPlatformFromUrl(url)
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading row {row.RowNumber()}: {ex.Message}");
                }
            }

            return data;
        }

        private async Task<List<string>> ValidateListingsAsync(List<UploadedProductData> listings)
        {
            var errors = new List<string>();
            if (listings == null || listings.Count == 0)
            {
                errors.Add("No valid listings found in the file");
                return errors;
            }

            var duplicateIds = listings
                .Select(l => l.ListingId?.Trim())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .GroupBy(id => id)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToList();

            if (duplicateIds.Count > 0)
            {
                errors.Add($"Duplicate Listing IDs found: {string.Join(", ", duplicateIds)}");
            }

            var ids = listings
                .Select(l => l.ListingId?.Trim())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct()
                .ToArray();

            if (ids.Length > 0)
            {
                try
                {
                    var response = await ArtemisGlobalClient.Instance.ValidateListingsApi
                        .Execute(ids)
                        .ConfigureAwait(false);
                    var responseBody = await response.Content.ReadAsStringAsync()
                        .ConfigureAwait(false);

                    using JsonDocument doc = JsonDocument.Parse(responseBody);
                    JsonElement root = doc.RootElement;

                    if (root.TryGetProperty("error", out JsonElement errorElement))
                    {
                        string errorMessage = errorElement.GetString() ?? "Validation error";
                        if (root.TryGetProperty("missing_ids", out JsonElement missingIdsElement))
                        {
                            var missingIds = missingIdsElement.EnumerateArray()
                                .Select(id => id.ToString())
                                .ToList();

                            errors.Add(errorMessage);
                            errors.AddRange(missingIds);
                            return errors;
                        }
                        errors.Add(errorMessage);
                        return errors;
                    }
                }
                catch (HttpRequestException ex)
                {
                    errors.Add($"API request failed: {ex.Message}");
                    return errors;
                }
                catch (Exception ex)
                {
                    errors.Add($"An error occurred: {ex.Message}");
                    return errors;
                }
            }

            for (int i = 0; i < listings.Count; i++)
            {
                int rowNumber = i + 2;
                var listingId = listings[i].ListingId?.Trim();

                if (string.IsNullOrWhiteSpace(listingId))
                {
                    errors.Add($"row {rowNumber}: Listing ID cannot be empty");
                }
                else if (!long.TryParse(listingId, out var parsedId) || parsedId <= 0)
                {
                    errors.Add($"row {rowNumber}: Listing ID must be positive (Value: {listingId})");
                }

                if (string.IsNullOrWhiteSpace(listings[i].ProductUrl))
                {
                    errors.Add($"row {rowNumber}: Product URL cannot be empty");
                }
                else
                {
                    if (!listings[i].ProductUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                        !listings[i].ProductUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    {
                        listings[i].ProductUrl = "https://" + listings[i].ProductUrl;
                    }

                    if (!Uri.TryCreate(listings[i].ProductUrl, UriKind.Absolute, out Uri uriResult) ||
                        (uriResult.Scheme != Uri.UriSchemeHttp && uriResult.Scheme != Uri.UriSchemeHttps))
                    {
                        errors.Add($"row {rowNumber}: Invalid URL format - {listings[i].ProductUrl}");
                    }
                }
            }

            return errors;
        }

        public static string GetPlatformFromUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return "Other|Not Supported";
            }

            var normalizedUrl = url.Trim().ToLowerInvariant();
            if (!normalizedUrl.StartsWith("http://") && !normalizedUrl.StartsWith("https://"))
            {
                normalizedUrl = $"https://{normalizedUrl}";
            }

            if (!Uri.TryCreate(normalizedUrl, UriKind.Absolute, out var uri))
            {
                return "Other|Not Supported";
            }

            var host = uri.Host.ToLowerInvariant();
            var platforms = GetPlatformsCache();

            foreach (var platform in platforms)
            {
                var domains = platform.Domains;
                if (domains == null)
                {
                    continue;
                }

                if (domains.Any(domain => host.Contains(domain.ToLowerInvariant())))
                {
                    if (platform.Availability == 1 && !string.IsNullOrWhiteSpace(platform.Name))
                    {
                        return platform.Name;
                    }

                    return string.IsNullOrWhiteSpace(platform.Name)
                        ? "Other|Not Supported"
                        : $"{platform.Name}|Not Supported";
                }
            }

            return "Other|Not Supported";
        }

        private static List<CrawlerPlatform> GetPlatformsCache()
        {
            if (_platformsCache != null)
            {
                return _platformsCache;
            }

            try
            {
                using var context = new ProductCheckerDbContext();
                _platformsCache = context.CrawlerPlatforms.ToList();
            }
            catch
            {
                _platformsCache = new List<CrawlerPlatform>();
            }

            return _platformsCache;
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ClearPreviewData();
            ClearSelections();
            UpdateSelectionSummary();
            _customFileName = string.Empty;
        }

        private void UploadClearButton_Click(object sender, RoutedEventArgs e)
        {
            ClearPreviewData();
            ClearUploadFileInfo();
        }

        private void ClearAllData()
        {
            _currentFilePath = string.Empty;
            _customFileName = string.Empty;
            ClearPreviewData();
            ClearSelections();
            UpdateSelectionSummary();
            ClearUploadFileInfo();
            UpdateStatusBar("Ready");
        }

        private void ClearPreviewData()
        {
            DataGridPreview.ItemsSource = null;
            UpdateDataGridVisibility(false);
            UpdateStatsDisplay();
            _uploadedData.Clear();
            StartButton.IsEnabled = false;
            RecordCountText.Text = "0 records loaded";
        }

        private void ClearUploadFileInfo()
        {
            if (UploadFilePathText != null)
            {
                UploadFilePathText.Text = "No file selected";
            }
            if (UploadFileInfoItems != null)
            {
                UploadFileInfoItems.ItemsSource = null;
            }
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new string(value.Where(c => !invalid.Contains(c)).ToArray()).Trim();

            return sanitized;
        }

        private void UpdateStatsDisplay()
        {
            // This method is intentionally left empty as it was only updating a variable that wasn't used
        }

        private void UpdateDataGridVisibility(bool hasData)
        {
            if (hasData)
            {
                EmptyState.Visibility = Visibility.Collapsed;
                DataGridPreview.Visibility = Visibility.Visible;
            }
            else
            {
                EmptyState.Visibility = Visibility.Visible;
                DataGridPreview.Visibility = Visibility.Collapsed;
            }
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (_uploadedData.Count == 0)
            {
                MessageBox.Show("No data to process. Please preview listings first.",
                              "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (FilterModeRadio?.IsChecked == true)
            {
                ShowCustomFileNameModal();
                return;
            }

            _customFileName = string.Empty;
            ShowConfirmProcessingModal();
        }

        private void ShowConfirmProcessingModal()
        {
            //ModalRecordCount.Text = $"{_uploadedData.Count} records";
            //ModalFileName.Text = string.IsNullOrEmpty(_currentFilePath) || FilePathText.Text == "No file selected" ?
            //    "Unknown file" : Path.GetFileName(_currentFilePath);
            //ModalUserName.Text = Environment.UserName ?? "SystemUser";

            ShowModal(ConfirmProcessingModal);
        }

        private void ShowCustomFileNameModal()
        {
            if (ModalCustomFileNameTextBox != null)
            {
                ModalCustomFileNameTextBox.Text = string.Empty;
                ModalCustomFileNameTextBox.Focus();
            }

            ShowModal(FileNameModal);
        }

        private void ShowModal(Border modal)
        {
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
            };

            modal.BeginAnimation(OpacityProperty, animation);
        }

        private void CancelProcessingButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(ConfirmProcessingModal);
        }

        private void CancelFileNameButton_Click(object sender, RoutedEventArgs e)
        {
            _customFileName = string.Empty;
            HideModal(FileNameModal);
        }

        private void ConfirmFileNameButton_Click(object sender, RoutedEventArgs e)
        {
            _customFileName = string.Empty;

            if (ModalCustomFileNameTextBox != null)
            {
                var raw = ModalCustomFileNameTextBox.Text?.Trim();
                if (!string.IsNullOrWhiteSpace(raw))
                {
                    _customFileName = SanitizeFileName(raw);
                }
            }

            FileNameModal.Visibility = Visibility.Collapsed;
            ShowConfirmProcessingModal();
        }

        private void ConfirmProcessingButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(ConfirmProcessingModal);
            StartProcessing();
        }

        private void CloseErrorButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(ErrorModal);
        }

        private void NewRequestButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(SuccessModal);
            ClearAllData();
            UpdateStatusBar("Ready for next request");
        }

        private void ViewRequestsSuccessButton_Click(object sender, RoutedEventArgs e)
        {
            HideModal(SuccessModal);
            ViewRequestsButton_Click(sender, e);
        }

        private void StartProcessing()
        {
            ProgressText.Text = "Creating request...";
            ProgressSubText.Text = "Initializing database connection...";
            MainProgressBar.Value = 0;
            MainProgressBar.IsIndeterminate = false;
            ProgressOverlay.Visibility = Visibility.Visible;
            StartButton.IsEnabled = false;
            BrowseButton.IsEnabled = false;
            ClearButton.IsEnabled = false;

            _processingWorker.RunWorkerAsync(_uploadedData);
        }

        private void ProcessingWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var data = e.Argument as List<UploadedProductData>;
            var result = new ProcessingResult();

            try
            {
                // Store the current progress message
                string currentMessage = "Connecting to database...";
                _processingWorker.ReportProgress(10, currentMessage);

                // Get file name
                string fileName;
                if (!string.IsNullOrWhiteSpace(_customFileName))
                {
                    fileName = _customFileName;
                }
                else if (!string.IsNullOrEmpty(_currentFilePath))
                {
                    fileName = Path.GetFileName(_currentFilePath);
                }
                else
                {
                    fileName = $"QuerySelection-{DateTime.Now:yyyyMMdd_HHmmss}";
                }

                using var context = new ProductCheckerDbContext();

                // Create new requestInfo
                var requestInfo = new RequestInfo
                {
                    User = Environment.UserName ?? "SystemUser",
                    FileName = fileName,
                    Environment = ConfigurationManager.GetEnvironment(),
                    CreatedAt = DateTime.Now
                };

                // Add and save to get ID
                context.RequestInfos.Add(requestInfo);
                context.SaveChanges();

                // Create new request
                var request = new Request
                {
                    RequestInfoId = requestInfo.Id,
                    Status = RequestStatus.PENDING,
                    CreatedAt = DateTime.Now
                };

                currentMessage = "Creating request entry...";
                _processingWorker.ReportProgress(30, currentMessage);
                context.Requests.Add(request);
                context.SaveChanges();

                _currentRequestId = request.Id;
                _currentRecordsProcessed = data.Count;
                result.RequestId = request.Id;
                result.RecordsProcessed = data.Count;
                currentMessage = $"Request #{request.Id} created. Adding product listings...";
                _processingWorker.ReportProgress(50, currentMessage);

                // Insert listings
                int totalRecords = data.Count;
                int processedRecords = 0;
                int batchSize = 100;

                for (int i = 0; i < totalRecords; i += batchSize)
                {
                    var batch = data.Skip(i).Take(batchSize).ToList();

                    foreach (var item in batch)
                    {
                        var listing = new ProductListing
                        {
                            RequestInfoId = requestInfo.Id,
                            ListingId = item.ListingId,
                            CaseNumber = item.CaseNumber,
                            Url = item.ProductUrl,
                            Platform = item.Platform,
                            CreatedAt = DateTime.Now
                        };

                        context.ProductListings.Add(listing);
                        processedRecords++;
                    }

                    int progress = 50 + (int)((double)processedRecords / totalRecords * 40);
                    currentMessage = $"Processing records... ({processedRecords}/{totalRecords})";
                    _processingWorker.ReportProgress(progress, currentMessage);

                    context.SaveChanges();
                    context.ChangeTracker.Clear();
                }

                currentMessage = "Finalizing request...";
                _processingWorker.ReportProgress(95, currentMessage);

                result.Success = true;
                result.RecordsProcessed = totalRecords;
                result.Message = $"Request #{request.Id} created with {totalRecords} listings";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
            }

            e.Result = result;
        }

        private void ProcessingWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                MainProgressBar.Value = e.ProgressPercentage;

                if (e.UserState is string message)
                {
                    ProgressText.Text = message;

                    // Update subtext based on progress
                    if (e.ProgressPercentage < 30)
                    {
                        ProgressSubText.Text = "Establishing database connection...";
                    }
                    else if (e.ProgressPercentage < 50)
                    {
                        ProgressSubText.Text = "Saving request information...";
                    }
                    else if (e.ProgressPercentage < 90)
                    {
                        ProgressSubText.Text = "Adding product listings to database...";
                    }
                    else
                    {
                        ProgressSubText.Text = "Completing the request...";
                    }
                }
            });
        }

        private void ProcessingWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressOverlay.Visibility = Visibility.Collapsed;
                MainProgressBar.IsIndeterminate = false;
                StartButton.IsEnabled = true;
                BrowseButton.IsEnabled = true;
                ClearButton.IsEnabled = true;

                if (e.Error != null)
                {
                    ShowErrorModal($"Error: {e.Error.Message}", e.Error.StackTrace);
                    return;
                }

                var result = e.Result as ProcessingResult;

                if (result.Success)
                {
                    ShowSuccessModal(result.RequestId, result.RecordsProcessed);
                }
                else
                {
                    ShowErrorModal("Failed to process request", result.ErrorMessage);
                }
            });
        }

        private void ShowSuccessModal(int requestId, int recordsProcessed)
        {
            ShowModal(SuccessModal);
        }

        private void ShowErrorModal(string title, string details)
        {
            ErrorMessage.Text = title;
            ErrorDetails.Text = details ?? "No additional details available.";

            ShowModal(ErrorModal);
        }

                        private void UpdateStatusBar(string message)
        {
            var environment = ConfigurationManager.GetEnvironment();
            this.Title = $"{_applicationName} - {message} ({environment})";
        }

        private void ShowValidationOverlay(string title, string subText)
        {
            if (_isFileValidationActive)
            {
                return;
            }

            _isFileValidationActive = true;
            ProgressText.Text = title;
            ProgressSubText.Text = subText;
            MainProgressBar.IsIndeterminate = true;
            MainProgressBar.Value = 0;
            ProgressOverlay.Visibility = Visibility.Visible;
            BrowseButton.IsEnabled = false;
            ClearButton.IsEnabled = false;
            StartButton.IsEnabled = false;
        }

        private void HideValidationOverlay()
        {
            if (!_isFileValidationActive)
            {
                return;
            }

            _isFileValidationActive = false;
            ProgressOverlay.Visibility = Visibility.Collapsed;
            MainProgressBar.IsIndeterminate = false;
        }

        private void DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (!e.PropertyName.Equals("ListingId") &&
                !e.PropertyName.Equals("CaseNumber") &&
                !e.PropertyName.Equals("ProductUrl") &&
                !e.PropertyName.Equals("Platform"))
            {
                e.Cancel = true;
            }
        }

        private void ViewRequestsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ShowViewRequestsPage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening requests view: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NavUploadButton_Click(object sender, RoutedEventArgs e)
        {
            ShowUploadPage();
        }

        private void NavRequestsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowViewRequestsPage();
        }

        private void ShowUploadPage()
        {
            MainContentFrame.Content = null;
            MainContentFrame.Visibility = Visibility.Collapsed;
            UploadPageRoot.Visibility = Visibility.Visible;
            UpdateNavSelection(false);
        }

        private void ShowViewRequestsPage()
        {
            UploadPageRoot.Visibility = Visibility.Collapsed;
            MainContentFrame.Visibility = Visibility.Visible;
            MainContentFrame.Navigate(new ViewRequestsPage());
            UpdateNavSelection(true);
        }

        private void UpdateNavSelection(bool isRequestsPage)
        {
            if (NavUploadButton != null)
            {
                NavUploadButton.IsChecked = !isRequestsPage;
            }

            if (NavRequestsButton != null)
            {
                NavRequestsButton.IsChecked = isRequestsPage;
            }
        }

        private class ProcessingResult
        {
            public bool Success { get; set; }
            public int RequestId { get; set; }
            public int RecordsProcessed { get; set; }
            public string Message { get; set; }
            public string ErrorMessage { get; set; }
        }

        private class FilterOption : INotifyPropertyChanged
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

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private class SelectedFilterItem
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
    }

    public class UploadedProductData
    {
        public string ListingId { get; set; }
        public string CaseNumber { get; set; }
        public string ProductUrl { get; set; }
        public string Platform { get; set; }
    }
}















