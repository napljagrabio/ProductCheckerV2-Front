using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using ProductCheckerV2.Common;
using ProductCheckerV2.Database;
using ProductCheckerV2.Database.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Animation;

namespace ProductCheckerV2
{
    public partial class MainWindow : Window
    {
        private List<UploadedProductData> _uploadedData = new List<UploadedProductData>();
        private BackgroundWorker _processingWorker;
        private string _applicationName = "Product Checker V2";
        private bool _isDragOver = false;
        private string _currentFilePath = string.Empty;
        private int _currentRequestId = 0;
        private int _currentRecordsProcessed = 0;

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
                this.Title = _applicationName;
                LogoImage.Source = new BitmapImage(new Uri("Assets/Logo.ico".AbsPath()));
                InitializeDatabase();
                UpdateStatusBar("Ready - Select or drag & drop Excel file");
                UpdateStatsDisplay();
                UpdateDataGridVisibility(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Application initialization failed:\n\n{ex.Message}",
                    "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void SetupDragDrop()
        {
            this.AllowDrop = true;
            this.PreviewDragOver += MainWindow_PreviewDragOver;
            this.PreviewDrop += MainWindow_PreviewDrop;

            // Add drag/drop visual states
            UploadArea.PreviewDragOver += UploadArea_PreviewDragOver;
            UploadArea.PreviewDragLeave += UploadArea_PreviewDragLeave;
            UploadArea.PreviewDrop += UploadArea_PreviewDrop;
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

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
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
                ValidateAndLoadFile(dialog.FileName);
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

        private void MainWindow_PreviewDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    ValidateAndLoadFile(files[0]);
                }
            }
        }

        private void UploadArea_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (!_isDragOver && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                _isDragOver = true;
                UploadArea.Background = new SolidColorBrush(
                    Color.FromArgb(30, 67, 97, 238)); // Primary color with opacity
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
        }

        private void UploadArea_PreviewDragLeave(object sender, DragEventArgs e)
        {
            _isDragOver = false;
            UploadArea.Background = Brushes.Transparent;
        }

        private void UploadArea_PreviewDrop(object sender, DragEventArgs e)
        {
            _isDragOver = false;
            UploadArea.Background = Brushes.Transparent;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    ValidateAndLoadFile(files[0]);
                }
            }
            e.Handled = true;
        }

        private void ValidateAndLoadFile(string filePath)
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

                LoadExcelFile(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading file:\n\n{ex.Message}",
                    "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadExcelFile(string filePath)
        {
            try
            {
                _currentFilePath = filePath;
                FilePathText.Text = Path.GetFileName(filePath);
                UpdateStatusBar($"Loading file: {Path.GetFileName(filePath)}...");

                // Read Excel data
                var data = ReadExcelData(filePath);
                _uploadedData = data;

                if (_uploadedData.Count == 0)
                {
                    MessageBox.Show("No valid data found in the Excel file.\n\n" +
                                  "Please ensure the file contains columns:\n" +
                                  "A: Listing ID, B: Case Number, C: URL",
                        "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearAllData();
                    return;
                }

                // Update UI
                DataGridPreview.ItemsSource = _uploadedData;
                UpdateDataGridVisibility(true);
                UpdateStatsDisplay();

                // Update record count text
                RecordCountText.Text = $"{_uploadedData.Count} records loaded";

                // Update file info items
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
                FileInfoItems.ItemsSource = infoItems;

                StartButton.IsEnabled = true;
                UpdateStatusBar($"Ready - {_uploadedData.Count} records loaded");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel file:\n\n{ex.Message}",
                    "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearAllData();
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

        private string GetPlatformFromUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                return "UNKNOWN";

            url = url.ToLower();

            if (url.Contains("amazon.")) return "AMAZON";
            if (url.Contains("ebay.")) return "EBAY";
            if (url.Contains("walmart.")) return "WALMART";
            if (url.Contains("target.")) return "TARGET";
            if (url.Contains("bestbuy.")) return "BESTBUY";

            return "OTHER";
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ClearAllData();
        }

        private void ClearAllData()
        {
            _currentFilePath = string.Empty;
            FilePathText.Text = "No file selected";
            DataGridPreview.ItemsSource = null;
            UpdateDataGridVisibility(false);
            UpdateStatsDisplay();
            _uploadedData.Clear();
            StartButton.IsEnabled = false;
            UpdateStatusBar("Ready");

            // Clear file info items
            FileInfoItems.ItemsSource = null;
            RecordCountText.Text = "0 records loaded";
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
                MessageBox.Show("No data to process. Please load an Excel file first.",
                              "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

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
                string fileName = string.IsNullOrEmpty(_currentFilePath) ?
                    "Unknown file" : Path.GetFileName(_currentFilePath);

                using var context = new ProductCheckerDbContext();

                // Create new requestInfo
                var requestInfo = new RequestInfos
                {
                    User = Environment.UserName ?? "SystemUser",
                    FileName = fileName,
                    CreatedAt = DateTime.Now
                };

                // Add and save to get ID
                context.RequestInfos.Add(requestInfo);
                context.SaveChanges();

                // Create new request
                var request = new Requests
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
                        var listing = new ProductListings
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
            SuccessRequestId.Text = $"Request #{requestId} created";
            SuccessRecordCount.Text = $"{recordsProcessed} records processed";
            SuccessUserName.Text = Environment.UserName ?? "SystemUser";
            SuccessTimestamp.Text = DateTime.Now.ToString("hh:mm tt");

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
            this.Title = $"{_applicationName} - {message}";
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
                var viewWindow = new ViewRequestsWindow();
                viewWindow.Owner = this;
                viewWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening requests view: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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