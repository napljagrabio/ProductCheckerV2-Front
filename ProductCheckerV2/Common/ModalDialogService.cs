using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace ProductCheckerV2.Common
{
    public enum ModalDialogType
    {
        Info,
        Success,
        Warning,
        Error
    }

    public static class ModalDialogService
    {
        private static Window? _bannerWindow;
        private static DispatcherTimer? _bannerTimer;
        private static bool _isConnectionLostBannerVisible;

        public static void Show(string message, string title, MessageBoxButton buttons, MessageBoxImage image)
        {
            _ = buttons;
            Show(message, title, MapType(image), Application.Current?.MainWindow);
        }

        public static void Show(string message, string title, MessageBoxButton buttons, MessageBoxImage image, Window? owner)
        {
            _ = buttons;
            Show(message, title, MapType(image), owner);
        }

        public static void Show(string message, string title, ModalDialogType type = ModalDialogType.Info, Window? owner = null)
        {
            var resolvedOwner = owner ?? Application.Current?.MainWindow;
            var canUseOwner = resolvedOwner != null && resolvedOwner.IsLoaded;

            if (!canUseOwner && resolvedOwner != null)
            {
                void OnLoaded(object? sender, RoutedEventArgs e)
                {
                    resolvedOwner.Loaded -= OnLoaded;
                    Show(message, title, type, resolvedOwner);
                }

                resolvedOwner.Loaded += OnLoaded;
                return;
            }

            var palette = GetPalette(type);
            try
            {
                var dialog = new Window
                {
                    Title = title,
                    Width = 460,
                    SizeToContent = SizeToContent.Height,
                    MaxHeight = 640,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStartupLocation = canUseOwner ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen,
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = Brushes.Transparent,
                    ShowInTaskbar = false
                };

                if (canUseOwner)
                {
                    dialog.Owner = resolvedOwner;
                }

                var chromeBorder = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(12),
                    Effect = new System.Windows.Media.Effects.DropShadowEffect
                    {
                        BlurRadius = 18,
                        ShadowDepth = 0,
                        Opacity = 0.2,
                        Color = Color.FromRgb(15, 23, 42)
                    }
                };

                var root = new Grid();
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var header = new Border
                {
                    Background = palette.HeaderBackground,
                    CornerRadius = new CornerRadius(12, 12, 0, 0),
                    Padding = new Thickness(18, 14, 18, 14)
                };

                var headerText = new TextBlock
                {
                    Text = title,
                    FontSize = 16,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = palette.HeaderForeground,
                    VerticalAlignment = VerticalAlignment.Center
                };
                header.Child = headerText;
                Grid.SetRow(header, 0);
                root.Children.Add(header);

                var body = new TextBlock
                {
                    Text = message,
                    Margin = new Thickness(0),
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                    TextWrapping = TextWrapping.Wrap
                };

                var bodyScroll = new ScrollViewer
                {
                    Margin = new Thickness(18, 16, 18, 8),
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    MaxHeight = 340,
                    Content = body
                };
                Grid.SetRow(bodyScroll, 1);
                root.Children.Add(bodyScroll);

                var footer = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(18, 8, 18, 16)
                };

                var okButton = new Button
                {
                    Content = "OK",
                    MinWidth = 86,
                    Height = 34,
                    Padding = new Thickness(16, 6, 16, 6),
                    Background = palette.ButtonBackground,
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                okButton.Click += (_, _) => dialog.Close();
                footer.Children.Add(okButton);

                Grid.SetRow(footer, 2);
                root.Children.Add(footer);

                chromeBorder.Child = root;
                dialog.Content = chromeBorder;

                dialog.ShowDialog();
            }
            catch
            {
                if (Application.Current?.Dispatcher != null && !Application.Current.Dispatcher.HasShutdownStarted)
                {
                    Application.Current.Dispatcher.BeginInvoke(
                        DispatcherPriority.Background,
                        new Action(() => { }));
                }
            }
        }

        public static void ShowBanner(string message, ModalDialogType type = ModalDialogType.Info, int autoHideMilliseconds = 0, Window? owner = null)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
            {
                return;
            }

            if (!dispatcher.CheckAccess())
            {
                dispatcher.Invoke(() => ShowBanner(message, type, autoHideMilliseconds, owner));
                return;
            }

            var resolvedOwner = owner ?? Application.Current?.MainWindow;
            var canUseOwner = resolvedOwner != null && resolvedOwner.IsLoaded;
            var palette = GetPalette(type);

            CloseBanner();

            var banner = new Window
            {
                Width = 560,
                Height = 60,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = Brushes.Transparent,
                ShowInTaskbar = false,
                Topmost = true,
                ShowActivated = false,
                WindowStartupLocation = WindowStartupLocation.Manual
            };

            if (canUseOwner)
            {
                banner.Owner = resolvedOwner;
            }

            var root = new Border
            {
                Background = palette.HeaderBackground,
                BorderBrush = palette.ButtonBackground,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(14, 10, 14, 10),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius = 16,
                    ShadowDepth = 0,
                    Opacity = 0.24,
                    Color = Color.FromRgb(15, 23, 42)
                }
            };

            var text = new TextBlock
            {
                Text = message,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = palette.HeaderForeground,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            root.Child = text;
            banner.Content = root;

            banner.Loaded += (_, _) =>
            {
                PositionBanner(banner, resolvedOwner);
            };

            _bannerWindow = banner;
            banner.Show();

            if (autoHideMilliseconds > 0)
            {
                _bannerTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(autoHideMilliseconds)
                };
                _bannerTimer.Tick += (_, _) =>
                {
                    _bannerTimer?.Stop();
                    _bannerTimer = null;
                    CloseBanner();
                };
                _bannerTimer.Start();
            }
        }

        public static void ShowConnectionLostBanner(Window? owner = null)
        {
            if (_isConnectionLostBannerVisible)
            {
                return;
            }

            _isConnectionLostBannerVisible = true;
            ShowBanner(
                "Connection lost. Showing saved data until internet is restored.",
                ModalDialogType.Warning,
                autoHideMilliseconds: 0,
                owner: owner);
        }

        public static void ShowConnectionRestoredBanner(Window? owner = null)
        {
            if (!_isConnectionLostBannerVisible)
            {
                return;
            }

            _isConnectionLostBannerVisible = false;
            ShowBanner(
                "Connection restored. Live updates resumed.",
                ModalDialogType.Success,
                autoHideMilliseconds: 3000,
                owner: owner);
        }

        public static void CloseBanner()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
            {
                return;
            }

            if (!dispatcher.CheckAccess())
            {
                dispatcher.Invoke(CloseBanner);
                return;
            }

            _bannerTimer?.Stop();
            _bannerTimer = null;

            if (_bannerWindow == null)
            {
                return;
            }

            try
            {
                _bannerWindow.Close();
            }
            catch
            {
                // Ignore close errors.
            }
            finally
            {
                _bannerWindow = null;
            }
        }

        private static void PositionBanner(Window banner, Window? owner)
        {
            if (banner == null)
            {
                return;
            }

            const double topMargin = 24;
            if (owner != null && owner.IsVisible)
            {
                banner.Left = owner.Left + Math.Max(12, (owner.ActualWidth - banner.Width) / 2);
                banner.Top = owner.Top + topMargin;
                return;
            }

            var area = SystemParameters.WorkArea;
            banner.Left = area.Left + Math.Max(12, (area.Width - banner.Width) / 2);
            banner.Top = area.Top + topMargin;
        }

        private static ModalDialogType MapType(MessageBoxImage image)
        {
            return image switch
            {
                MessageBoxImage.Error => ModalDialogType.Error,
                MessageBoxImage.Warning => ModalDialogType.Warning,
                MessageBoxImage.Information => ModalDialogType.Info,
                _ => ModalDialogType.Info
            };
        }

        private static (SolidColorBrush HeaderBackground, SolidColorBrush HeaderForeground, SolidColorBrush ButtonBackground) GetPalette(ModalDialogType type)
        {
            return type switch
            {
                ModalDialogType.Success => (
                    new SolidColorBrush(Color.FromRgb(236, 253, 245)),
                    new SolidColorBrush(Color.FromRgb(5, 150, 105)),
                    new SolidColorBrush(Color.FromRgb(16, 185, 129))),
                ModalDialogType.Warning => (
                    new SolidColorBrush(Color.FromRgb(255, 251, 235)),
                    new SolidColorBrush(Color.FromRgb(217, 119, 6)),
                    new SolidColorBrush(Color.FromRgb(245, 158, 11))),
                ModalDialogType.Error => (
                    new SolidColorBrush(Color.FromRgb(254, 242, 242)),
                    new SolidColorBrush(Color.FromRgb(220, 38, 38)),
                    new SolidColorBrush(Color.FromRgb(239, 68, 68))),
                _ => (
                    new SolidColorBrush(Color.FromRgb(239, 246, 255)),
                    new SolidColorBrush(Color.FromRgb(37, 99, 235)),
                    new SolidColorBrush(Color.FromRgb(59, 130, 246)))
            };
        }
    }
}
