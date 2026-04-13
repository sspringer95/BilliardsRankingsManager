using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using SkiaSharp;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Xceed.Wpf.Toolkit;
using MessageBox = System.Windows.MessageBox;



namespace BilliardsRankingsManager
{
    // ============================================
    // CENTRALIZED STYLING CONFIGURATION
    // ============================================
    public static class StyleConfig
    {
        // Export Colors (Felt Theme)
        public static SKColor BackgroundColor = new SKColor(27, 94, 32);      // Dark green felt
        public static SKColor TitleBackgroundColor = new SKColor(21, 71, 24); // Darker green
        public static SKColor TextColor = SKColor.Parse("#FFFFFF");           // White
        public static SKColor RankNumberColor = SKColor.Parse("#FFD700");     // Gold
        public static SKColor BorderColor = new SKColor(139, 69, 19);         // Saddle brown (wood)

        // Export Layout Defaults
        public static int DefaultWidth = 800;
        public static int DefaultHeight = 1200;
        public static int DefaultFontSize = 28;
        public static int DefaultSpacing = 45;
        public static int DefaultPadding = 40;
        public static int DefaultTitleFontSize = 48;
        public static string DefaultFontFamily = "Arial";

        // UI Colors
        public static Brush UIBackground = new SolidColorBrush(Color.FromRgb(240, 240, 240));

        public static Brush UILightBrown = new SolidColorBrush(Color.FromRgb(201, 181, 156));
        public static Brush UILighterBrown = new SolidColorBrush(Color.FromRgb(217, 207, 199));
        public static Brush ExpandedBackground = new SolidColorBrush(Color.FromRgb(225, 225, 225)); 
    }

    // ============================================
    // RANK FORMAT OPTIONS
    // ============================================
    public enum RankFormat
    {
        NumberDot,          // "1. Player Name"
        HashDash,           // "#1 - Player Name"
        PlaceColon          // "1st Place: Player Name"
    }

    public static class RankFormatHelper
    {
        public static string GetDisplayName(RankFormat format)
        {
            return format switch
            {
                RankFormat.NumberDot => "1. Player Name",
                RankFormat.HashDash => "#1 - Player Name",
                RankFormat.PlaceColon => "1st Place: Player Name",
                _ => "1. Player Name"
            };
        }

        public static string FormatRank(int rank, RankFormat format)
        {
            return format switch
            {
                RankFormat.NumberDot => $"{rank}.",
                RankFormat.HashDash => $"#{rank} -",
                RankFormat.PlaceColon => $"{GetOrdinal(rank)} Place:",
                _ => $"{rank}."
            };
        }

        private static string GetOrdinal(int number)
        {
            if (number <= 0) return number.ToString();

            switch (number % 100)
            {
                case 11:
                case 12:
                case 13:
                    return number + "th";
            }

            switch (number % 10)
            {
                case 1: return number + "st";
                case 2: return number + "nd";
                case 3: return number + "rd";
                default: return number + "th";
            }
        }
    }

    // ============================================
    // DATA MODELS
    // ============================================
    public class RankingEntry : INotifyPropertyChanged
    {
        private string _playerName;
        public int Rank { get; set; }

        public string PlayerName
        {
            get => _playerName;
            set
            {
                if (_playerName != value)
                {
                    _playerName = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class AppSettings : INotifyPropertyChanged
    {
        private string _exportPath = "";
        private string _logoPath = "";
        private int _exportWidth = StyleConfig.DefaultWidth;
        private int _exportHeight = StyleConfig.DefaultHeight;
        private int _fontSize = StyleConfig.DefaultFontSize;
        private int _spacing = StyleConfig.DefaultSpacing;
        private int _paddingTop = 40;
        private int _paddingBottom = 40;
        private int _paddingLeft = 40;
        private int _paddingRight = 40;
        private int _logoHeight = 100;
        private int _logoTitleSpacing = 20;
        private string _titleText = "TOP 20 RANKINGS";
        private string _fontFamily = StyleConfig.DefaultFontFamily;
        private RankFormat _rankFormat = RankFormat.NumberDot;

        private int[] _customColors = new int[16]
        {
            0x1B5E20, // Forest Green (dark felt green)
            0x2E7D32, // Medium Green
            0x4CAF50, // Bright Green
            0x81C784, // Light Green
            0x8B4513, // Saddle Brown (wood)
            0xA0522D, // Sienna Brown
            0xD2691E, // Rich Brown
            0xF4A460, // Sandy Brown
            0x212121, // Near Black
            0x424242, // Dark Gray
            0x757575, // Medium Gray
            0xBDBDBD, // Light Gray
            0xFFD700, // Gold (ranking accent)
            0xB71C1C, // Deep Red
            0x1565C0, // Rich Blue
            0xF5F5F5  // Off White
        };

        public int[] CustomColors
        {
            get => _customColors;
            set { _customColors = value; OnPropertyChanged(); }
        }

        private double _splitterRatio = 1.5; // Default to 1.5*

        public double SplitterRatio
        {
            get => _splitterRatio;
            set { _splitterRatio = value; OnPropertyChanged(); }
        }


        private bool _settingsPanelVisible = true;

        public bool SettingsPanelVisible
        {
            get => _settingsPanelVisible;
            set { _settingsPanelVisible = value; OnPropertyChanged(); }
        }


        private int _playerCount = 20;
        private int _titleFontSize = 48;

        private bool _backgroundExpanded = false;
        private bool _textExpanded = false;
        private bool _spacingExpanded = false;

        public bool BackgroundExpanded
        {
            get => _backgroundExpanded;
            set { _backgroundExpanded = value; OnPropertyChanged(); }
        }

        public bool TextExpanded
        {
            get => _textExpanded;
            set { _textExpanded = value; OnPropertyChanged(); }
        }

        public bool SpacingExpanded
        {
            get => _spacingExpanded;
            set { _spacingExpanded = value; OnPropertyChanged(); }
        }


        public int TitleFontSize
        {
            get => _titleFontSize;
            set { _titleFontSize = value; OnPropertyChanged(); }
        }

        public int PlayerCount
        {
            get => _playerCount;
            set { _playerCount = value; OnPropertyChanged(); }
        }

      

        public string ExportPath
        {
            get => _exportPath;
            set { _exportPath = value; OnPropertyChanged(); }
        }

        public string LogoPath
        {
            get => _logoPath;
            set { _logoPath = value; OnPropertyChanged(); }
        }

        public int ExportWidth
        {
            get => _exportWidth;
            set { _exportWidth = value; OnPropertyChanged(); }
        }

        public int ExportHeight
        {
            get => _exportHeight;
            set { _exportHeight = value; OnPropertyChanged(); }
        }

        public int FontSize
        {
            get => _fontSize;
            set { _fontSize = value; OnPropertyChanged(); }
        }

        public int Spacing
        {
            get => _spacing;
            set { _spacing = value; OnPropertyChanged(); }
        }

        public int PaddingTop
        {
            get => _paddingTop;
            set { _paddingTop = value; OnPropertyChanged(); }
        }

        public int PaddingBottom
        {
            get => _paddingBottom;
            set { _paddingBottom = value; OnPropertyChanged(); }
        }

        public int PaddingLeft
        {
            get => _paddingLeft;
            set { _paddingLeft = value; OnPropertyChanged(); }
        }

        public int PaddingRight
        {
            get => _paddingRight;
            set { _paddingRight = value; OnPropertyChanged(); }
        }

        public int LogoHeight
        {
            get => _logoHeight;
            set { _logoHeight = value; OnPropertyChanged(); }
        }

        public int LogoTitleSpacing
        {
            get => _logoTitleSpacing;
            set { _logoTitleSpacing = value; OnPropertyChanged(); }
        }

        public string TitleText
        {
            get => _titleText;
            set { _titleText = value; OnPropertyChanged(); }
        }

        public string FontFamily
        {
            get => _fontFamily;
            set { _fontFamily = value; OnPropertyChanged(); }
        }

        public RankFormat RankFormat
        {
            get => _rankFormat;
            set { _rankFormat = value; OnPropertyChanged(); }
        }

        private string _backgroundColor = "#1B5E20"; // Dark green
        private string _borderColor = "#8B4513"; // Brown
        private string _titleTextColor = "#FFFFFF"; // White
        private string _playerTextColor = "#FFFFFF"; // White
        private string _rankNumberColor = "#FFD700"; // Gold
        private bool _hasTextOutline = false;
        private string _outlineColor = "#000000"; // Black
        private int _outlineWidth = 2;

        public string BackgroundColor
        {
            get => _backgroundColor;
            set { _backgroundColor = value; OnPropertyChanged(); }
        }

        public string BorderColor
        {
            get => _borderColor;
            set { _borderColor = value; OnPropertyChanged(); }
        }

        public string TitleTextColor
        {
            get => _titleTextColor;
            set { _titleTextColor = value; OnPropertyChanged(); }
        }

        public string PlayerTextColor
        {
            get => _playerTextColor;
            set { _playerTextColor = value; OnPropertyChanged(); }
        }

        public string RankNumberColor
        {
            get => _rankNumberColor;
            set { _rankNumberColor = value; OnPropertyChanged(); }
        }

        public bool HasTextOutline
        {
            get => _hasTextOutline;
            set { _hasTextOutline = value; OnPropertyChanged(); }
        }

        public string OutlineColor
        {
            get => _outlineColor;
            set { _outlineColor = value; OnPropertyChanged(); }
        }

        public int OutlineWidth
        {
            get => _outlineWidth;
            set { _outlineWidth = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class RankingsData
    {
        public List<RankingEntry> Rankings { get; set; } = new List<RankingEntry>();
    }

    // ============================================
    // MAIN APPLICATION
    // ============================================
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly string _dataPath;
        private readonly string _settingsPath;
        private System.Timers.Timer _saveTimer;
        private bool _isDirty = false;

        private bool _settingsDirty = false;
        private string _currentPresetName = null;
        private ComboBox _presetComboBox = null; // Store reference to update it

        // Add these three new fields:
        private Window _previewWindow = null;
        private Image _previewImage = null;
        private System.Timers.Timer _previewUpdateTimer;

        public ObservableCollection<RankingEntry> Rankings { get; set; }
        public AppSettings Settings { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            // Setup data paths
            var appData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BilliardsRankingsManager"
            );
            Directory.CreateDirectory(appData);
            _dataPath = Path.Combine(appData, "rankings.json");
            _settingsPath = Path.Combine(appData, "settings.json");

            // Initialize data
            Rankings = new ObservableCollection<RankingEntry>();
            Settings = new AppSettings();

            LoadData();
            LoadSettings();

            // Setup auto-save timer (debounced)
            _saveTimer = new System.Timers.Timer(500);
            _saveTimer.Elapsed += (s, e) =>
            {
                _saveTimer.Stop();
                Dispatcher.Invoke(() => SaveData());
            };

            InitializeComponent();

            // Subscribe to property changes
            foreach (var ranking in Rankings)
            {
                ranking.PropertyChanged += (s, e) => MarkDirty();
            }

            // ADD THIS NEW CODE:
            // Subscribe to Rankings changes for live preview
            foreach (var ranking in Rankings)
            {
                ranking.PropertyChanged += (s, e) => TriggerPreviewUpdate();
            }

            // Subscribe to Settings changes for live preview
            Settings.PropertyChanged += (s, e) => TriggerPreviewUpdate();
        }


        private Grid _mainGrid;
        private MenuItem _toggleSettingsMenuItem;

        // Stores the "Star" ratio (e.g., 1.5*) so we don't lose it when hiding
        private GridLength _lastSettingsColumnWidth = new GridLength(1.5, GridUnitType.Star);

        // Stores the actual pixel width we removed so we can add it back exactly
        private double _collapsedPanelWidth = 450;


        private void InitializeComponent()
        {
            Title = "Billiards Top 20 Rankings Manager";
            Width = 800;
            Height = 700;
            Background = StyleConfig.UIBackground;

            // Create main container
            var dockPanel = new DockPanel();

            // Create menu bar
            var menuBar = new Menu
            {
                Background = StyleConfig.UIBackground, // Match window color
                IsMainMenu = true,
                Padding = new Thickness(0), // Remove container padding
                VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(menuBar, Dock.Top);

            // File Menu
            var fileMenu = new MenuItem
            {
                Header = "_File",
                Height = 25, // FORCE SMALLER HEIGHT (Standard Desktop is ~22-25)
                Padding = new Thickness(10, 0, 10, 0), // Reduce internal spacing
                VerticalAlignment = VerticalAlignment.Center
            };

            var importMenuItem = new MenuItem { Header = "_Import Rankings from CSV..." };
            importMenuItem.Click += ImportRankingsFromCSV_Click;

            var exportMenuItem = new MenuItem { Header = "_Export Rankings to CSV..." };
            exportMenuItem.Click += ExportRankingsToCSV_Click;

            fileMenu.Items.Add(importMenuItem);
            fileMenu.Items.Add(exportMenuItem);
            fileMenu.Items.Add(new Separator());

            var setExportLocationItem = new MenuItem { Header = "Set _Export Location..." };
            setExportLocationItem.Click += SetExportLocation_Click;

            var browseLogoItem = new MenuItem { Header = "Browse for _Logo..." };
            browseLogoItem.Click += BrowseLogo_Click;

            fileMenu.Items.Add(setExportLocationItem);
            fileMenu.Items.Add(browseLogoItem);
            fileMenu.Items.Add(new Separator());

            var loadPresetItem = new MenuItem { Header = "_Load Preset..." };
            loadPresetItem.Click += (s, e) =>
            {
                if (_settingsDirty)
                {
                    var result = MessageBox.Show(
                        "Save current settings as a preset before loading?",
                        "Unsaved Changes",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                        SavePresetDialog();
                    else if (result == MessageBoxResult.Cancel)
                        return;
                }

                LoadPresetDialog();
            };

            var savePresetItem = new MenuItem { Header = "_Save Preset..." };
            savePresetItem.Click += (s, e) => SavePresetDialog();

            fileMenu.Items.Add(loadPresetItem);
            fileMenu.Items.Add(savePresetItem);

            fileMenu.Items.Add(new Separator());

            var resetColorsItem = new MenuItem { Header = "Reset Custom _Colors to Default..." };
            resetColorsItem.Click += (s, e) => ResetCustomColorsToDefault();
            fileMenu.Items.Add(resetColorsItem);


            // Settings Panel Toggle Menu Item
            var toggleSettingsItem = new MenuItem
            {
                Header = Settings.SettingsPanelVisible ? "_Hide Settings Panel" : "_Show Settings Panel",
                Height = 25, // MATCH FILE MENU HEIGHT
                Padding = new Thickness(10, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            _toggleSettingsMenuItem = toggleSettingsItem;
            toggleSettingsItem.Click += (s, e) => ToggleSettingsPanel();

            menuBar.Items.Add(fileMenu);
            menuBar.Items.Add(toggleSettingsItem);

            dockPanel.Children.Add(menuBar);

            // ---------------------------------------------------------
            // REPLACEMENT BLOCK FOR GRID SETUP
            // ---------------------------------------------------------

            // Create main grid
            var mainGrid = new Grid();
            _mainGrid = mainGrid; // Store reference

            // Column 0: Rankings (Takes available space)
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition
            {
                Width = new GridLength(1, GridUnitType.Star)
            });

            // Column 1: The Splitter (Fixed width lane)
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition
            {
                Width = new GridLength(5) // Fixed pixel width for the handle
            });

            // Column 2: Settings (Takes available space based on ratio)
            // We use the initial ratio from settings here
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition
            {
                Width = new GridLength(Settings.SplitterRatio, GridUnitType.Star)
            });

            // --- 1. Left Panel (Rankings) ---
            var leftPanel = CreateRankingsPanel();
            Grid.SetColumn(leftPanel, 0);
            mainGrid.Children.Add(leftPanel);

            // --- 2. The Grid Splitter ---
            var gridSplitter = new GridSplitter
            {
                Width = 5, // Matches Column 1 width
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(Color.FromRgb(200, 200, 200)), // Slightly lighter for visibility
                ShowsPreview = true, // Shows a ghost line while dragging (better performance)
                ResizeDirection = GridResizeDirection.Columns,
                ResizeBehavior = GridResizeBehavior.PreviousAndNext, // Resizes Col 0 and Col 2
                Cursor = Cursors.SizeWE
            };
            Grid.SetColumn(gridSplitter, 1); // Sits in its own dedicated column
            mainGrid.Children.Add(gridSplitter);

            // --- 3. Right Panel (Settings) ---
            var rightPanel = CreateSettingsPanel();
            Grid.SetColumn(rightPanel, 2);
            mainGrid.Children.Add(rightPanel);

            // ---------------------------------------------------------

            dockPanel.Children.Add(mainGrid);
            Content = dockPanel;

     
        }

        private void ToggleSettingsPanel()
        {
            if (_mainGrid == null || _toggleSettingsMenuItem == null) return;

            Settings.SettingsPanelVisible = !Settings.SettingsPanelVisible;
            SaveSettings();

            // Columns: 0=Rankings, 1=Splitter, 2=Settings
            var splitterCol = _mainGrid.ColumnDefinitions[1];
            var settingsCol = _mainGrid.ColumnDefinitions[2];

            if (Settings.SettingsPanelVisible)
            {
                // --- SHOWING ---

                // 1. Restore the layout
                // Restore splitter to fixed size
                splitterCol.Width = new GridLength(5);
                // Restore settings to the last known ratio (Memory!)
                settingsCol.Width = _lastSettingsColumnWidth;

                _toggleSettingsMenuItem.Header = "_Hide Settings Panel";

                // 2. Resize the Window (only if not maximized)
                if (WindowState == System.Windows.WindowState.Normal)
                {
                    // Add back the exact pixels we removed last time
                    Width += _collapsedPanelWidth;
                }
            }
            else
            {
                // --- HIDING ---

                // 1. Capture State (Memory!)
                // Save the current ratio so we don't reset to default later
                _lastSettingsColumnWidth = settingsCol.Width;

                // Calculate total pixels to remove (Splitter + Settings Panel)
                double totalWidthToRemove = splitterCol.ActualWidth + settingsCol.ActualWidth;

                // Save this value so we can restore the window size later
                // (Guard against 0 to prevent shrinking window to nothing on rapid toggles)
                if (totalWidthToRemove > 10)
                {
                    _collapsedPanelWidth = totalWidthToRemove;
                }

                // 2. Collapse Columns
                splitterCol.Width = new GridLength(0);
                settingsCol.Width = new GridLength(0);

                _toggleSettingsMenuItem.Header = "_Show Settings Panel";

                // 3. Resize the Window (only if not maximized)
                if (WindowState == System.Windows.WindowState.Normal)
                {
                    Width -= _collapsedPanelWidth;
                }
            }
        }

        private void ResetCustomColorsToDefault()
        {
            var result = MessageBox.Show(
                "This will reset the custom color palette in the color picker to the default.\n\n" +
                "Your current custom colors will be lost. Continue?",
                "Reset Custom Colors",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                Settings.CustomColors = new int[16]
                {
                    0x1B5E20, // Forest Green (dark felt green)
                    0x2E7D32, // Medium Green
                    0x4CAF50, // Bright Green
                    0x81C784, // Light Green
                    0x8B4513, // Saddle Brown (wood)
                    0xA0522D, // Sienna Brown
                    0xD2691E, // Rich Brown
                    0xF4A460, // Sandy Brown
                    0x212121, // Near Black
                    0x424242, // Dark Gray
                    0x757575, // Medium Gray
                    0xBDBDBD, // Light Gray
                    0xFFD700, // Gold (ranking accent)
                    0xB71C1C, // Deep Red
                    0x1565C0, // Rich Blue
                    0xF5F5F5  // Off White
                };

                SaveSettings();


            }
        }

        private void ExportRankingsToCSV_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "CSV Files|*.csv",
                Title = "Export Rankings to CSV",
                FileName = "rankings.csv"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var lines = new List<string>();

                    // Export all 50 rankings (empty or not)
                    for (int i = 0; i < 50; i++)
                    {
                        var name = Rankings[i].PlayerName ?? "";

                        // Wrap in quotes if contains comma
                        if (name.Contains(","))
                        {
                            name = $"\"{name}\"";
                        }

                        lines.Add(name);
                    }

                    File.WriteAllLines(dialog.FileName, lines);

                    MessageBox.Show("Rankings exported successfully!", "Export Complete",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting rankings: {ex.Message}", "Export Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ImportRankingsFromCSV_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv",
                Title = "Import Rankings from CSV"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var lines = File.ReadAllLines(dialog.FileName);

                    // Import up to 50 lines
                    for (int i = 0; i < 50; i++)
                    {
                        if (i < lines.Length)
                        {
                            var name = lines[i].Trim();

                            // Remove quotes if present
                            if (name.StartsWith("\"") && name.EndsWith("\""))
                            {
                                name = name.Substring(1, name.Length - 2);
                            }

                            Rankings[i].PlayerName = name;
                        }
                        else
                        {
                            // CSV has fewer than 50 lines - clear remaining
                            Rankings[i].PlayerName = "";
                        }
                    }

                    // Show warning if CSV had more than 50 lines
                    if (lines.Length > 50)
                    {
                        MessageBox.Show($"CSV contained {lines.Length} entries. Only the first 50 were imported.",
                            "Import Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show($"Successfully imported {lines.Length} rankings!", "Import Complete",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error importing rankings: {ex.Message}", "Import Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SwapRankings(int index1, int index2)
        {
            if (index1 < 0 || index1 >= 50 || index2 < 0 || index2 >= 50)
                return;

            var temp = Rankings[index1].PlayerName;
            Rankings[index1].PlayerName = Rankings[index2].PlayerName;
            Rankings[index2].PlayerName = temp;
        }

        private System.Windows.Threading.DispatcherTimer _refreshTimer;

        private void ScheduleRefresh()
        {
            // Cancel existing timer if any
            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer = null;
            }

            // Create new timer for 150ms debounce
            _refreshTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(150)
            };

            _refreshTimer.Tick += (s, e) =>
            {
                _refreshTimer.Stop();
                _refreshTimer = null;
                RefreshPreview(); // Calls the existing RefreshPreview method
            };

            _refreshTimer.Start();
        }




        private List<Button> _swapButtons = new List<Button>();
        private int? _pendingSwapIndex = null;
        private TextBlock _swapBanner;

        private ScrollViewer CreateRankingsPanel()
        {
            var stack = new StackPanel { Margin = new Thickness(15) };

            // Create error popup
            var errorPopup = new System.Windows.Controls.Primitives.Popup
            {
                Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom,
                StaysOpen = false,
                Opacity = 0.8

            };

            var errorText = new TextBlock
            {
                Text = "Please enter a number between 1 and 50",
                Background = Brushes.LightYellow,
                Padding = new Thickness(2),
                FontSize = 16
            };

            var errorBorder = new Border
            {
                Child = errorText,
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Background = Brushes.LightYellow
            };

            errorPopup.Child = errorBorder;

            var titleFontSize = 15;

            // Title with player count input
            var titlePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };

            titlePanel.Children.Add(new TextBlock
            {
                Text = "List of  ",
                FontSize = titleFontSize,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            });

            var countBox = CreateSelectAllTextBox();
            countBox.Text = Settings.PlayerCount.ToString();
            countBox.Width = 40;
            countBox.FontSize = 18;
            countBox.FontWeight = FontWeights.Bold;
            countBox.Padding = new Thickness(5, 2, 5, 2);
            countBox.TextAlignment = TextAlignment.Center;

            errorPopup.PlacementTarget = countBox;

            countBox.TextChanged += (s, e) =>
            {
                if (int.TryParse(countBox.Text, out int value))
                {
                    if (value >= 1 && value <= 50)
                    {
                        Settings.PlayerCount = value;
                        SaveSettings();
                        UpdateVisibleRankings();
                        errorPopup.IsOpen = false;
                    }
                    else
                    {
                        countBox.Text = "";
                        errorPopup.IsOpen = true;

                        // Auto-hide after 3 seconds
                        var timer = new System.Windows.Threading.DispatcherTimer
                        {
                            Interval = TimeSpan.FromSeconds(2)
                        };
                        timer.Tick += (t, args) =>
                        {
                            errorPopup.IsOpen = false;
                            timer.Stop();
                        };
                        timer.Start();
                    }
                }
            };
            titlePanel.Children.Add(countBox);

            titlePanel.Children.Add(new TextBlock
            {
                Text = " players",
                FontSize = titleFontSize,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            });

            stack.Children.Add(titlePanel);

            _swapBanner = new TextBlock
            {
                Text = "Click ⇅ Select next to a player to begin a swap.",
                FontSize = 12,
                FontStyle = FontStyles.Italic,
                Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 4, 8, 4),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8),
                Height = 40
            };
            stack.Children.Add(_swapBanner);

            // Clear previous entries
            _rankingEntryGrids.Clear();
            _swapButtons.Clear();

            // Create 50 ranking entries
            for (int i = 0; i < 50; i++)
            {
                var entryGrid = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                entryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                entryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                entryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var rankLabel = new TextBlock
                {
                    Text = $"{i + 1}.",
                    Width = 30,
                    FontSize = 13,
                    FontWeight = FontWeights.Bold,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(rankLabel, 0);

                var nameBox = CreateSelectAllTextBox();
                nameBox.FontSize = 13;
                nameBox.Padding = new Thickness(5);
                nameBox.Margin = new Thickness(5, 0, 5, 0);
                nameBox.SetBinding(TextBox.TextProperty, new System.Windows.Data.Binding("PlayerName")
                {
                    Source = Rankings[i],
                    Mode = System.Windows.Data.BindingMode.TwoWay,
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                });
                Grid.SetColumn(nameBox, 1);

                int index = i;

                var swapBtn = new Button
                {
                    Content = "⇅ Select",
                    Width = 85,
                    Height = 22,
                    FontSize = 10,
                    Padding = new Thickness(0),
                    IsTabStop = false
                };
                swapBtn.Click += (s, e) => HandleSwapClick(index);

                _swapButtons.Add(swapBtn);
                Grid.SetColumn(swapBtn, 2);

                entryGrid.Children.Add(rankLabel);
                entryGrid.Children.Add(nameBox);
                entryGrid.Children.Add(swapBtn);

                // Set initial visibility
                entryGrid.Visibility = i < Settings.PlayerCount ? Visibility.Visible : Visibility.Collapsed;

                _rankingEntryGrids.Add(entryGrid);
                stack.Children.Add(entryGrid);
            }

            return new ScrollViewer
            {
                Content = stack,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
        }

        private List<Grid> _rankingEntryGrids = new List<Grid>();
        private void UpdateVisibleRankings()
        {
            // Cancel any pending swap when the list size changes
            _pendingSwapIndex = null;
            RefreshSwapState();

            for (int i = 0; i < _rankingEntryGrids.Count; i++)
            {
                _rankingEntryGrids[i].Visibility = i < Settings.PlayerCount
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }
        }

        private void HandleSwapClick(int index)
        {
            if (_pendingSwapIndex == null)
            {
                // First click: select this player
                _pendingSwapIndex = index;
            }
            else if (_pendingSwapIndex == index)
            {
                // Same player clicked again: cancel
                _pendingSwapIndex = null;
            }
            else
            {
                // Second click on a different player: execute swap
                SwapRankings(_pendingSwapIndex.Value, index);
                _pendingSwapIndex = null;
            }

            RefreshSwapState();
        }

        private void RefreshSwapState()
        {
            bool pending = _pendingSwapIndex.HasValue;

            if (pending)
            {
                var name = Rankings[_pendingSwapIndex!.Value].PlayerName;
                var display = string.IsNullOrWhiteSpace(name) ? $"#{_pendingSwapIndex.Value + 1}" : name;
                _swapBanner.Text = $"Swapping {display} — click ↔ Swap next to another player, or cancel.";
                _swapBanner.Foreground = new SolidColorBrush(Color.FromRgb(100, 60, 0));
                _swapBanner.Background = new SolidColorBrush(Color.FromRgb(255, 236, 179));
            }
            else
            {
                _swapBanner.Text = "Click ⇅ Select next to a player to begin a swap.";
                _swapBanner.Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120));
                _swapBanner.Background = Brushes.Transparent;
            }

            for (int i = 0; i < _swapButtons.Count; i++)
            {
                var btn = _swapButtons[i];
                var grid = _rankingEntryGrids[i];

                if (!pending)
                {
                    btn.Content = "⇅ Select";
                    grid.Background = Brushes.Transparent;
                }
                else if (i == _pendingSwapIndex.Value)
                {
                    btn.Content = "✕ Cancel";
                    grid.Background = new SolidColorBrush(Color.FromRgb(255, 236, 179));
                }
                else
                {
                    btn.Content = "↔ Swap";
                    grid.Background = Brushes.Transparent;
                }
            }
        }

        private Separator CreateDivider()
        {
            return new Separator
            {
                Margin = new Thickness(0, 20, 0, 15),
                Background = new SolidColorBrush(Color.FromRgb(200, 200, 200))
            };
        }

        private void LoadPresetDialog()
        {
            var presetsFolder = GetPresetsFolder();

            var dialog = new OpenFileDialog
            {
                Filter = "JSON Files|*.json",
                Title = "Load Settings Preset",
                InitialDirectory = presetsFolder
            };

            if (dialog.ShowDialog() == true)
            {
                var presetName = Path.GetFileNameWithoutExtension(dialog.FileName);
                LoadPresetFromFile(presetName);
            }
        }

        private List<string> GetCommonFonts()
        {
            return new List<string>
                {
                    "Arial",
                    "Arial Black",
                    "Calibri",
                    "Cambria",
                    "Comic Sans MS",
                    "Consolas",
                    "Courier New",
                    "Georgia",
                    "Impact",
                    "Lucida Console",
                    "Lucida Sans Unicode",
                    "Microsoft Sans Serif",
                    "Palatino Linotype",
                    "Segoe UI",
                    "Tahoma",
                    "Times New Roman",
                    "Trebuchet MS",
                    "Verdana",
                    "Roboto",
                    "Open Sans",
                    "Montserrat",
                    "Lato",
                    "Oswald",
                    "Raleway",
                    "PT Sans"
                };
        }

        private ScrollViewer CreateSettingsPanel()
        {


            var stack = new StackPanel { Margin = new Thickness(15) };

            // Settings Title
            var title = new TextBlock
            {
                Text = "Settings",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 0)
            };

            stack.Children.Add(title);

            stack.Children.Add(CreateDivider());

            var previewBtn = CreateButton("Preview", (s, e) => _ = ShowPreview(), isPrimary: false);
            previewBtn.Name = "PreviewButton";
            stack.Children.Add(previewBtn);

            var pngBtn = CreateButton("Generate PNG", (s, e) => _ = ExportToPNG(), isPrimary: true);
            pngBtn.Margin = new Thickness(0, 10, 0, 0);
            stack.Children.Add(pngBtn);

            stack.Children.Add(CreateDivider());

            // Logo Section
            stack.Children.Add(CreateSliderSetting("Logo Size:", Settings.LogoHeight, 50, 250,
                v => { Settings.LogoHeight = v; SaveSettings(); }));


            stack.Children.Add(CreateDivider());


            // Background Section (collapsible)
            var backgroundExpander = new Expander
            {
                Header = "Background (Size & Colors)",
                IsExpanded = Settings.BackgroundExpanded,
                Margin = new Thickness(0, 5, 0, 5),
                FontWeight = FontWeights.SemiBold,
                FontSize = 13,
                Background = Brushes.Transparent // Keep header transparent
            };

            var backgroundStack = new StackPanel();

            backgroundStack.Children.Add(CreateSliderSetting("Background Width (px):", Settings.ExportWidth, 100, 1920,
                v => { Settings.ExportWidth = v; SaveSettings(); }, 5));
            backgroundStack.Children.Add(CreateSliderSetting("Background Height (px):", Settings.ExportHeight, 100, 1920,
                v => { Settings.ExportHeight = v; SaveSettings(); }, 5));
            backgroundStack.Children.Add(CreateColorSetting("Background:",
                () => Settings.BackgroundColor,
                c => { Settings.BackgroundColor = c; SaveSettings(); ScheduleRefresh(); }));
            backgroundStack.Children.Add(CreateColorSetting("Border:",
                () => Settings.BorderColor,  // Func that reads current value
                c => { Settings.BorderColor = c; SaveSettings(); ScheduleRefresh(); }));


            // --- MATERIAL DESIGN REFACTOR START ---

            // 1. Create a Border to handle the Padding (Space inside the card)
            var backgroundBorder = new Border
            {
                Child = backgroundStack,
                Padding = new Thickness(16) // Adjust this number to increase/decrease internal space
            };

            // 2. Wrap that Border in a Material Design Card
            // The Card provides the "Paper" background color and the shadow.
            var backgroundCard = new MaterialDesignThemes.Wpf.Card
            {
                Content = backgroundBorder,
                Margin = new Thickness(8, 10, 8, 8), // Margin prevents the shadow from being clipped
                UniformCornerRadius = 6 // Optional: Rounds the corners slightly
            };

            // 3. Logic: Only show the Card when Expanded
            // Set initial state
            backgroundExpander.Content = Settings.BackgroundExpanded ? backgroundCard : null;

            // Handle Expand (Show Card)
            backgroundExpander.Expanded += (s, e) =>
            {
                Settings.BackgroundExpanded = true;
                SaveSettings();
                backgroundExpander.Content = backgroundCard;
            };

            // Handle Collapse (Hide Card/Shadow completely)
            backgroundExpander.Collapsed += (s, e) =>
            {
                Settings.BackgroundExpanded = false;
                SaveSettings();
                backgroundExpander.Content = null;
            };

            // --- MATERIAL DESIGN REFACTOR END ---

            stack.Children.Add(backgroundExpander);

            stack.Children.Add(CreateDivider());





            // Text Section (collapsible)
            var textExpander = new Expander
            {
                Header = "Text (Size, Color, and outline)",
                IsExpanded = Settings.TextExpanded,
                Margin = new Thickness(0, 5, 0, 5),
                FontWeight = FontWeights.SemiBold,
                FontSize = 13,
                Background = Brushes.Transparent
            };

            var textStack = new StackPanel();

            var textBorder = new Border
            {
                Child = textStack, // <--- IMPORTANT: This puts the Text settings inside
                Padding = new Thickness(16)
            };

            var textCard = new MaterialDesignThemes.Wpf.Card
            {
                Content = textBorder,
                Margin = new Thickness(8, 15, 8, 8),
                UniformCornerRadius = 6
            };

            // --- 1. Title Text Input ---
            var titleTextPanel = new StackPanel { Margin = new Thickness(0, 3, 0, 3) };
            titleTextPanel.Children.Add(new TextBlock { Text = "Title Text:", FontWeight = FontWeights.SemiBold, FontSize = 12 });
            var titleBox = new TextBox
            {
                Text = Settings.TitleText,
                Margin = new Thickness(0, 3, 0, 0),
                Padding = new Thickness(5)
            };

            titleBox.TextChanged += (s, e) =>
            {
                Settings.TitleText = titleBox.Text;
                SaveSettings();
            };
            titleTextPanel.Children.Add(titleBox);
            textStack.Children.Add(titleTextPanel);

            // --- 2. Sliders & Colors ---
            textStack.Children.Add(CreateSliderSetting("Title Font Size:", Settings.TitleFontSize, 24, 72,
                v => { Settings.TitleFontSize = v; SaveSettings(); }));
            textStack.Children.Add(CreateColorSetting("Title Color:",
            () => Settings.TitleTextColor,
            c => { Settings.TitleTextColor = c; SaveSettings(); ScheduleRefresh(); }));

            textStack.Children.Add(CreateSliderSetting("Player Name Font Size:", Settings.FontSize, 12, 48,
                v => { Settings.FontSize = v; SaveSettings(); }));
            textStack.Children.Add(CreateColorSetting("Player Text Color:",
                () => Settings.PlayerTextColor,
                c => { Settings.PlayerTextColor = c; SaveSettings(); ScheduleRefresh(); }));
            textStack.Children.Add(CreateColorSetting("Player Rank Color:",
                () => Settings.RankNumberColor,
                c => { Settings.RankNumberColor = c; SaveSettings(); ScheduleRefresh(); }));


            // --- 3. Font Family Dropdown ---
            var fontPanel = new StackPanel { Margin = new Thickness(0, 3, 0, 3) };
            fontPanel.Children.Add(new TextBlock { Text = "Font Family:", FontWeight = FontWeights.SemiBold, FontSize = 12 });

            var fontCombo = new ComboBox
            {
                Margin = new Thickness(0, 3, 0, 0),
                Padding = new Thickness(5),
                MaxDropDownHeight = 300
            };

            var commonFonts = GetCommonFonts();
            foreach (var fontName in commonFonts)
            {
                var item = new ComboBoxItem
                {
                    Content = fontName,
                    FontFamily = new FontFamily(fontName),
                    FontSize = 14
                };
                fontCombo.Items.Add(item);

                // Select current font
                if (fontName.Equals(Settings.FontFamily, StringComparison.OrdinalIgnoreCase))
                {
                    fontCombo.SelectedItem = item;
                }
            }

            // If current font not in list, add it and select it
            if (fontCombo.SelectedItem == null)
            {
                var customItem = new ComboBoxItem
                {
                    Content = Settings.FontFamily,
                    FontFamily = new FontFamily(Settings.FontFamily),
                    FontSize = 14
                };
                fontCombo.Items.Insert(0, customItem);
                fontCombo.SelectedItem = customItem;
            }

            fontCombo.SelectionChanged += (s, e) =>
            {
                if (fontCombo.SelectedItem is ComboBoxItem selected)
                {
                    Settings.FontFamily = selected.Content.ToString();
                    SaveSettings();
                }
            };

            fontPanel.Children.Add(fontCombo);
            textStack.Children.Add(fontPanel);

            // --- 4. Rank Format Dropdown ---
            var rankFormatPanel = new StackPanel { Margin = new Thickness(0, 3, 0, 3) };
            rankFormatPanel.Children.Add(new TextBlock { Text = "Rank Format:", FontWeight = FontWeights.SemiBold, FontSize = 12 });

            var rankFormatCombo = new ComboBox
            {
                Margin = new Thickness(0, 3, 0, 0),
                Padding = new Thickness(5)
            };

            foreach (RankFormat format in Enum.GetValues(typeof(RankFormat)))
            {
                rankFormatCombo.Items.Add(new ComboBoxItem
                {
                    Content = RankFormatHelper.GetDisplayName(format),
                    Tag = format
                });
            }

            for (int i = 0; i < rankFormatCombo.Items.Count; i++)
            {
                var item = rankFormatCombo.Items[i] as ComboBoxItem;
                if ((RankFormat)item.Tag == Settings.RankFormat)
                {
                    rankFormatCombo.SelectedIndex = i;
                    break;
                }
            }

            rankFormatCombo.SelectionChanged += (s, e) =>
            {
                if (rankFormatCombo.SelectedItem is ComboBoxItem selected)
                {
                    Settings.RankFormat = (RankFormat)selected.Tag;
                    SaveSettings();
                }
            };

            rankFormatPanel.Children.Add(rankFormatCombo);
            textStack.Children.Add(rankFormatPanel);

            // --- 5. Text Outline Logic ---
            var outlineCheckPanel = new StackPanel { Margin = new Thickness(0, 4, 0, 4) };
            var outlineCheck = new CheckBox
            {
                Content = "Enable Text Outline",
                IsChecked = Settings.HasTextOutline
            };
            outlineCheckPanel.Children.Add(outlineCheck);
            textStack.Children.Add(outlineCheckPanel);

            // Outline options (hidden by default)
            var outlineOptionsPanel = new StackPanel
            {
                Visibility = Settings.HasTextOutline ? Visibility.Visible : Visibility.Collapsed,
                Margin = new Thickness(15, 0, 0, 0) // Indented
            };

            outlineOptionsPanel.Children.Add(CreateColorSetting("Outline Color:",
                () => Settings.OutlineColor,
                c => { Settings.OutlineColor = c; SaveSettings(); ScheduleRefresh(); }));
            outlineOptionsPanel.Children.Add(CreateSliderSetting("Outline Width:", Settings.OutlineWidth, 1, 10,
                v => { Settings.OutlineWidth = v; SaveSettings(); }));

            outlineCheck.Checked += (s, e) =>
            {
                Settings.HasTextOutline = true;
                SaveSettings();
                outlineOptionsPanel.Visibility = Visibility.Visible;
            };
            outlineCheck.Unchecked += (s, e) =>
            {
                Settings.HasTextOutline = false;
                SaveSettings();
                outlineOptionsPanel.Visibility = Visibility.Collapsed;
            };

            textStack.Children.Add(outlineOptionsPanel);


            // --- MATERIAL DESIGN REFACTOR START ---

            

            // 3. Logic: Only show the Card when Expanded
            textExpander.Content = Settings.TextExpanded ? textCard : null;

            textExpander.Expanded += (s, e) =>
            {
                Settings.TextExpanded = true;
                SaveSettings();
                textExpander.Content = textCard;
            };

            textExpander.Collapsed += (s, e) =>
            {
                Settings.TextExpanded = false;
                SaveSettings();
                textExpander.Content = null;
            };

            // --- MATERIAL DESIGN REFACTOR END ---

            stack.Children.Add(textExpander);

            stack.Children.Add(CreateDivider());





            // Spacing & Margins Section (collapsible)
            var spacingExpander = new Expander
            {
                Header = "Spacing & Margins",
                IsExpanded = Settings.SpacingExpanded,
                Margin = new Thickness(0, 5, 0, 5),
                FontWeight = FontWeights.SemiBold,
                FontSize = 13,
                Background = Brushes.Transparent
            };

            var spacingStack = new StackPanel();


            var spacingBorder = new Border
            {
                Child = spacingStack, // <--- IMPORTANT: This puts the Text settings inside
                Padding = new Thickness(16)
            };

            var spacingCard = new MaterialDesignThemes.Wpf.Card
            {
                Content = spacingBorder,
                Margin = new Thickness(8, 15, 8, 8),
                UniformCornerRadius = 6
            };


            spacingStack.Children.Add(CreateSliderSetting("Line Spacing:", Settings.Spacing, 20, 80,
                v => { Settings.Spacing = v; SaveSettings(); }));
            spacingStack.Children.Add(CreateSliderSetting("Spacing between logo and title:", Settings.LogoTitleSpacing, 0, 100,
                v => { Settings.LogoTitleSpacing = v; SaveSettings(); }));
            spacingStack.Children.Add(CreateNumberSetting("Margin Top:", Settings.PaddingTop,
                v => { Settings.PaddingTop = v; SaveSettings(); }));
            spacingStack.Children.Add(CreateNumberSetting("Margin Bottom:", Settings.PaddingBottom,
                v => { Settings.PaddingBottom = v; SaveSettings(); }));
            spacingStack.Children.Add(CreateNumberSetting("Margin Left:", Settings.PaddingLeft,
                v => { Settings.PaddingLeft = v; SaveSettings(); }));
            spacingStack.Children.Add(CreateNumberSetting("Margin Right:", Settings.PaddingRight,
                v => { Settings.PaddingRight = v; SaveSettings(); }));

            

            // Logic
            spacingExpander.Content = Settings.SpacingExpanded ? spacingCard : null;

            spacingExpander.Expanded += (s, e) =>
            {
                Settings.SpacingExpanded = true;
                SaveSettings();
                spacingExpander.Content = spacingCard;
            };

            spacingExpander.Collapsed += (s, e) =>
            {
                Settings.SpacingExpanded = false;
                SaveSettings();
                spacingExpander.Content = null;
            };

            stack.Children.Add(spacingExpander);





            stack.Children.Add(new Border { Height = 20 });


            return new ScrollViewer
            {
                Content = stack,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible
            };
        }

        private TextBox CreateSelectAllTextBox()
        {
            var textBox = new TextBox
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(6, 4, 6, 4)
            };

            // Select all on focus
            textBox.GotKeyboardFocus += (s, e) => textBox.SelectAll();
            textBox.PreviewMouseLeftButtonDown += (s, e) =>
            {
                if (!textBox.IsKeyboardFocusWithin)
                {
                    e.Handled = true;
                    textBox.Focus();
                }
            };

            return textBox;
        }

        private StackPanel CreateNumberSetting(string label, int defaultValue, Action<int> onChange)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 3, 0, 3) };
            panel.Children.Add(new TextBlock { Text = label, FontWeight = FontWeights.SemiBold });

            var textBox = new TextBox
            {
                Text = defaultValue.ToString(),
                Margin = new Thickness(0, 5, 0, 0),
                Padding = new Thickness(5)
            };

            textBox.TextChanged += (s, e) =>
            {
                if (int.TryParse(textBox.Text, out int value) && value > 0)
                {
                    onChange(value);
                }
            };
            panel.Children.Add(textBox);
            return panel;
        }

        private StackPanel CreateSliderSetting(string label, int defaultValue, int min, int max, Action<int> onChange, int tick = 1)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 4, 0, 4) };

            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var labelText = new TextBlock
            {
                Text = label,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(labelText, 0);

            var valueText = new TextBlock
            {
                Text = defaultValue.ToString(),
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(10, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(valueText, 1);

            headerGrid.Children.Add(labelText);
            headerGrid.Children.Add(valueText);
            panel.Children.Add(headerGrid);

            var slider = new Slider
            {
                Minimum = min,
                Maximum = max,
                Value = defaultValue,
                Margin = new Thickness(0, 5, 0, 0),
                TickFrequency = tick,
                IsSnapToTickEnabled = true
            };

            slider.ValueChanged += (s, e) =>
            {
                int value = (int)slider.Value;
                valueText.Text = value.ToString();
                onChange(value);
            };

            panel.Children.Add(slider);
            return panel;
        }

        private StackPanel CreateColorSetting(string label, Func<string> getCurrentColor, Action<string> onChange)
        {
            // Use a Grid for the main layout to better control vertical alignment
            var panel = new StackPanel { Margin = new Thickness(0, 5, 0, 5), Orientation = Orientation.Horizontal };

            // 1. Main Label
            var labelText = new TextBlock
            {
                Text = label,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Width = 120
            };
            panel.Children.Add(labelText);

            // 2. THE PREVIEW BOX
            // We create the inner grid (Checkerboard + Color) first
            var previewContentGrid = new Grid();

            // Layer A: Checkerboard (Always visible)
            var checkerboard = new DrawingBrush
            {
                TileMode = TileMode.Tile,
                Viewport = new Rect(0, 0, 10, 10),
                ViewportUnits = BrushMappingMode.Absolute,
                Drawing = new GeometryDrawing
                {
                    Brush = Brushes.LightGray,
                    Geometry = new GeometryGroup
                    {
                        Children = new GeometryCollection
                {
                    new RectangleGeometry(new Rect(0, 0, 5, 5)),
                    new RectangleGeometry(new Rect(5, 5, 5, 5))
                }
                    }
                }
            };
            var checkerboardLayer = new Border { Background = checkerboard };
            previewContentGrid.Children.Add(checkerboardLayer);

            // Layer B: The Color Layer (Alpha changes here)
            var colorLayer = new Border();
            previewContentGrid.Children.Add(colorLayer);

            // The Container Border (Acts like the Button)
            var previewBox = new Border
            {
                Child = previewContentGrid,
                Width = 26, // Slightly wider, less tall
                Height = 26, // Standard button height
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(1), // JITTER FIX: Start with 1px padding
                CornerRadius = new CornerRadius(3),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 15, 0) // Space between box and slider
            };

            // 3. OPACITY SECTION
            // Using a Grid to place text exactly above the slider
            var opacityGrid = new Grid { VerticalAlignment = VerticalAlignment.Center };
            opacityGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Row 0: Text
            opacityGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Row 1: Slider

            // The Text (Top Right)
            var opacityTextLabel = new TextBlock
            {
                Text = "Opacity 100%",
                FontSize = 11,
                FontWeight = FontWeights.Bold, // "Dark Black" feel
                Foreground = Brushes.Black,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 2, 0) // Align with end of slider
            };
            Grid.SetRow(opacityTextLabel, 0);

            // The Slider (Bottom)
            var opacitySlider = new Slider
            {
                Minimum = 0,
                Maximum = 100,
                Value = 100,
                Width = 140,
                VerticalAlignment = VerticalAlignment.Top,
                IsSnapToTickEnabled = true,
                TickFrequency = 1
            };
            Grid.SetRow(opacitySlider, 1);

            opacityGrid.Children.Add(opacityTextLabel);
            opacityGrid.Children.Add(opacitySlider);


            // --- LOGIC ---

            Action updatePreview = () =>
            {
                try
                {
                    var currentColor = getCurrentColor();

                    if (currentColor.Equals("transparent", StringComparison.OrdinalIgnoreCase))
                    {
                        colorLayer.Background = Brushes.Transparent;
                        opacitySlider.Value = 0;
                        opacityTextLabel.Text = "Opacity 0%";
                    }
                    else
                    {
                        var color = (Color)ColorConverter.ConvertFromString(currentColor);

                        int opacityPercent = (int)Math.Round((color.A / 255.0) * 100);

                        // Only update slider if significantly different to prevent loop fighting
                        if (Math.Abs(opacitySlider.Value - opacityPercent) > 1)
                            opacitySlider.Value = opacityPercent;

                        opacityTextLabel.Text = $"Opacity {opacityPercent}%";
                        colorLayer.Background = new SolidColorBrush(color);
                    }
                }
                catch
                {
                    colorLayer.Background = Brushes.White;
                }
            };

            // Initialize
            updatePreview();

            // --- EVENTS ---

            // 1. JITTER-FREE HOVER EFFECT
            previewBox.MouseEnter += (s, e) =>
            {
                previewBox.BorderBrush = Brushes.Gray; // Darker border
                previewBox.BorderThickness = new Thickness(2); // Thicker line
                previewBox.Padding = new Thickness(0); // Remove padding so size stays constant
            };

            previewBox.MouseLeave += (s, e) =>
            {
                previewBox.BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180));
                previewBox.BorderThickness = new Thickness(1); // Thinner line
                previewBox.Padding = new Thickness(1); // Add padding back
            };

            // 2. Click Logic (WinForms Dialog)
            previewBox.MouseLeftButtonDown += (s, e) =>
            {
                var dialog = new System.Windows.Forms.ColorDialog
                {
                    FullOpen = true,
                    AnyColor = true,
                    CustomColors = Settings.CustomColors
                };

                try
                {
                    var currentColor = getCurrentColor();
                    if (!currentColor.Equals("transparent", StringComparison.OrdinalIgnoreCase))
                    {
                        var wpfColor = (Color)ColorConverter.ConvertFromString(currentColor);
                        dialog.Color = System.Drawing.Color.FromArgb(wpfColor.R, wpfColor.G, wpfColor.B);
                    }
                }
                catch { }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    byte alpha = (byte)Math.Round((opacitySlider.Value / 100.0) * 255);
                    var wpfColor = Color.FromArgb(alpha, dialog.Color.R, dialog.Color.G, dialog.Color.B);

                    onChange(wpfColor.ToString());
                    Settings.CustomColors = dialog.CustomColors;
                    SaveSettings();
                    updatePreview();
                }
            };

            // 3. Slider Logic
            opacitySlider.ValueChanged += (s, e) =>
            {
                int opacityPercent = (int)opacitySlider.Value;
                opacityTextLabel.Text = $"Opacity {opacityPercent}%";

                // Logic to update the actual color string
                try
                {
                    var currentColor = getCurrentColor();
                    byte alpha = (byte)Math.Round((opacityPercent / 100.0) * 255);

                    if (opacityPercent == 0)
                    {
                        onChange("Transparent");
                    }
                    else
                    {
                        // If currently transparent, assume White base, otherwise use current RGB
                        Color baseColor = Colors.White;

                        if (!currentColor.Equals("transparent", StringComparison.OrdinalIgnoreCase))
                        {
                            baseColor = (Color)ColorConverter.ConvertFromString(currentColor);
                        }

                        var newColor = Color.FromArgb(alpha, baseColor.R, baseColor.G, baseColor.B);
                        onChange(newColor.ToString());
                    }

                    // Visual update only (don't call full updatePreview to avoid recursion on slider)
                    colorLayer.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(getCurrentColor()));
                }
                catch { }
            };

            // Add to main panel
            panel.Children.Add(previewBox);
            panel.Children.Add(opacityGrid);

            return panel;
        }


        private Button CreateButton(string text, RoutedEventHandler handler, bool isPrimary = false)
        {
            var button = new Button
            {
                Content = text,
                Padding = new Thickness(15, 8, 15, 8),
                FontSize = 14,
                Cursor = Cursors.Hand
            };

            if (isPrimary)
            {
                button.Foreground = Brushes.White;
                button.FontWeight = FontWeights.Bold;
            }

            button.Click += handler;
            return button;
        }



        private void BrowseLogo_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp",
                Title = "Select Logo Image"
            };

            // Set initial directory to existing logo path if it exists
            if (!string.IsNullOrEmpty(Settings.LogoPath) && File.Exists(Settings.LogoPath))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(Settings.LogoPath);
            }

            if (dialog.ShowDialog() == true)
            {
                Settings.LogoPath = dialog.FileName;
                SaveSettings();
                UpdateStatusText("LogoStatus", Path.GetFileName(Settings.LogoPath));
                MessageBox.Show("Logo set successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SetExportLocation_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "PNG Image|*.png",
                Title = "Set Export Location",
                FileName = "top20_rankings.png"
            };

            // Set initial directory to existing export path if it exists
            if (!string.IsNullOrEmpty(Settings.ExportPath) && File.Exists(Settings.ExportPath))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(Settings.ExportPath);
                dialog.FileName = Path.GetFileName(Settings.ExportPath);
            }

            if (dialog.ShowDialog() == true)
            {
                Settings.ExportPath = dialog.FileName;
                SaveSettings();
                UpdateStatusText("LocationStatus", Settings.ExportPath);
                MessageBox.Show("Export location set successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void UpdateStatusText(string name, string text)
        {
            if (FindName(name) is TextBlock textBlock)
            {
                textBlock.Text = text;
            }
            else
            {
                // Find in visual tree
                FindAndUpdateTextBlock(Content as DependencyObject, name, text);
            }
        }

        private void FindAndUpdateTextBlock(DependencyObject parent, string name, string text)
        {
            if (parent == null) return;

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is TextBlock tb && tb.Name == name)
                {
                    tb.Text = text;
                    return;
                }

                FindAndUpdateTextBlock(child, name, text);
            }
        }

        private async Task ExportToPNG()
        {
            if (string.IsNullOrEmpty(Settings.ExportPath))
            {
                MessageBox.Show("Please set an export location first.", "Export Location Required",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var exportPath = Settings.ExportPath;
                if (!exportPath.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                {
                    exportPath = Path.ChangeExtension(exportPath, ".png");
                }

                await Task.Run(() => GenerateRankingsImage(exportPath));

                MessageBox.Show($"PNG exported successfully to:\n{exportPath}", "Export Complete",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting PNG: {ex.Message}", "Export Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private SKColor ParseColor(string colorString)
        {
            if (colorString.Equals("Transparent", StringComparison.OrdinalIgnoreCase))
            {
                return SKColors.Transparent;
            }

            try
            {
                var wpfColor = (Color)ColorConverter.ConvertFromString(colorString);
                return new SKColor(wpfColor.R, wpfColor.G, wpfColor.B, wpfColor.A);
            }
            catch
            {
                return SKColors.White;
            }
        }

        private byte[] GenerateRankingsImageToMemory()
        {
            var info = new SKImageInfo(Settings.ExportWidth, Settings.ExportHeight);

            using var surface = SKSurface.Create(info);
            var canvas = surface.Canvas;
            var bgColor = ParseColor(Settings.BackgroundColor);
            canvas.Clear(bgColor);

            // Draw border
            using var borderPaint = new SKPaint
            {
                Color = ParseColor(Settings.BorderColor),
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 8,
                IsAntialias = true
            };
            canvas.DrawRect(4, 4, info.Width - 8, info.Height - 8, borderPaint);

            int currentY = Settings.PaddingTop;

            int logoHeight = 0;

            // Draw Logo if available
            if (!string.IsNullOrEmpty(Settings.LogoPath) && File.Exists(Settings.LogoPath))
            {
                using var logoBitmap = SKBitmap.Decode(Settings.LogoPath);
                if (logoBitmap != null)
                {
                    logoHeight = Settings.LogoHeight;  // NEW
                    int logoWidth = (int)(logoBitmap.Width * (logoHeight / (float)logoBitmap.Height));
                    int logoX = (info.Width - logoWidth) / 2;

                    var logoRect = new SKRect(logoX, currentY, logoX + logoWidth, currentY + logoHeight);
                    canvas.DrawBitmap(logoBitmap, logoRect);
                    currentY += logoHeight + Settings.LogoTitleSpacing;  // NEW
                }
            }

            // Draw Title
            using var titlePaint = new SKPaint
            {
                Color = ParseColor(Settings.TitleTextColor),
                TextSize = Settings.TitleFontSize,
                IsAntialias = true,
                Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold),
                TextAlign = SKTextAlign.Center
            };

            if (Settings.HasTextOutline)
            {
                using var outlinePaint = new SKPaint
                {
                    Color = ParseColor(Settings.OutlineColor),
                    TextSize = Settings.TitleFontSize,
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold),
                    TextAlign = SKTextAlign.Center,
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = Settings.OutlineWidth
                };
                canvas.DrawText(Settings.TitleText, info.Width / 2, currentY, outlinePaint);
            }
            canvas.DrawText(Settings.TitleText, info.Width / 2, currentY, titlePaint);
            currentY += 30;

            // Draw line separator
            using var linePaint = new SKPaint
            {
                Color = ParseColor(Settings.RankNumberColor),
                StrokeWidth = 3,
                IsAntialias = true
            };
            int lineMargin = Settings.PaddingLeft;
            canvas.DrawLine(lineMargin, currentY, info.Width - lineMargin, currentY, linePaint);
            currentY += 30;

            // Get active rankings
            var activeRankings = Rankings.Where(r => r.Rank <= Settings.PlayerCount && !string.IsNullOrWhiteSpace(r.PlayerName))
                            .OrderBy(r => r.Rank)
                            .ToList();

            if (activeRankings.Count == 0)
            {
                using var emptyPaint = new SKPaint
                {
                    Color = StyleConfig.TextColor,
                    TextSize = Settings.FontSize,
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Italic),
                    TextAlign = SKTextAlign.Center
                };
                canvas.DrawText("No rankings available", info.Width / 2, currentY + 50, emptyPaint);
            }
            else
            {
                int availableHeight = info.Height - currentY - Settings.PaddingBottom;
                int requiredHeight = activeRankings.Count * Settings.Spacing;
                bool needsColumns = requiredHeight > availableHeight;

                if (needsColumns)
                {
                    DrawRankingsInColumns(canvas, activeRankings, currentY, info.Width, availableHeight);
                }
                else
                {
                    DrawRankingsSingleColumn(canvas, activeRankings, currentY, Settings.PaddingLeft + 20);
                }
            }

            // Return as byte array
            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            return data.ToArray();
        }

        private async Task ShowPreview()
        {
            try
            {
                // If preview window already exists and is open, just bring it to front
                if (_previewWindow != null && _previewWindow.IsVisible)
                {
                    _previewWindow.Activate();
                    return;
                }

                // Generate initial image
                byte[] imageBytes = await Task.Run(() => GenerateRankingsImageToMemory());

                // Create preview window
                _previewWindow = new Window
                {
                    Title = "Preview - Top 20 Rankings (Live)",
                    Width = 1000,
                    Height = 750,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
                };

                // Load image from memory
                var bitmap = new BitmapImage();
                using (var stream = new MemoryStream(imageBytes))
                {
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = stream;
                    bitmap.EndInit();
                }

                _previewImage = new Image
                {
                    Source = bitmap,
                    Stretch = Stretch.Uniform
                };

                var scrollViewer = new ScrollViewer
                {
                    Content = _previewImage,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                };

                _previewWindow.Content = scrollViewer;

                // Handle window closing
                _previewWindow.Closed += (s, e) =>
                {
                    _previewWindow = null;
                    _previewImage = null;
                    if (_previewUpdateTimer != null)
                    {
                        _previewUpdateTimer.Stop();
                    }

                    // Update button text when preview closes
                    UpdatePreviewButtonText("Preview");
                };

                // Update button text when preview opens
                UpdatePreviewButtonText("Preview Open (Live)");

                // Initialize the update timer (but don't start it yet)
                if (_previewUpdateTimer == null)
                {
                    _previewUpdateTimer = new System.Timers.Timer(150); // 150ms debounce
                    _previewUpdateTimer.Elapsed += async (s, e) =>
                    {
                        _previewUpdateTimer.Stop();
                        await Dispatcher.InvokeAsync(() => RefreshPreview());
                    };
                }

                _previewWindow.Show(); // Non-modal!
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating preview: {ex.Message}", "Preview Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdatePreviewButtonText(string text)
        {
            // Find the preview button in the visual tree and update its text
            if (Content is DependencyObject root)
            {
                var button = FindButtonByName(root, "PreviewButton");
                if (button != null)
                {
                    button.Content = text;
                }
            }
        }

        private Button FindButtonByName(DependencyObject parent, string name)
        {
            if (parent == null) return null;

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is Button btn && btn.Name == name)
                {
                    return btn;
                }

                var result = FindButtonByName(child, name);
                if (result != null) return result;
            }
            return null;
        }

        private async void RefreshPreview()
        {
            // Check if preview window still exists and is visible
            if (_previewWindow == null || !_previewWindow.IsVisible || _previewImage == null)
            {
                return;
            }

            try
            {
                // Generate new image in background
                byte[] imageBytes = await Task.Run(() => GenerateRankingsImageToMemory());

                // Load new image from memory
                var bitmap = new BitmapImage();
                using (var stream = new MemoryStream(imageBytes))
                {
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = stream;
                    bitmap.EndInit();
                }

                // Update the image source
                _previewImage.Source = bitmap;
            }
            catch (Exception ex)
            {
                // Silently fail - don't interrupt user experience with error popups
                // Could log this in production
                System.Diagnostics.Debug.WriteLine($"Preview refresh error: {ex.Message}");
            }
        }

        private void TriggerPreviewUpdate()
        {
            // Only update if preview window is open
            if (_previewWindow != null && _previewWindow.IsVisible && _previewUpdateTimer != null)
            {
                _previewUpdateTimer.Stop();
                _previewUpdateTimer.Start(); // Restart the 150ms countdown
            }
        }



        private void GenerateRankingsImage(string outputPath)
        {
            // Generate image bytes
            var imageBytes = GenerateRankingsImageToMemory();

            // Save to file
            File.WriteAllBytes(outputPath, imageBytes);
        }

        private void DrawRankingsSingleColumn(SKCanvas canvas, List<RankingEntry> rankings, int startY, int leftMargin)
        {
            using var rankPaint = new SKPaint
            {
                Color = ParseColor(Settings.RankNumberColor),
                TextSize = Settings.FontSize,
                IsAntialias = true,
                Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold)
            };

            using var namePaint = new SKPaint
            {
                Color = ParseColor(Settings.PlayerTextColor),
                TextSize = Settings.FontSize,
                IsAntialias = true,
                Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Normal)
            };

            SKPaint rankOutlinePaint = null;
            SKPaint nameOutlinePaint = null;

            if (Settings.HasTextOutline)
            {
                rankOutlinePaint = new SKPaint
                {
                    Color = ParseColor(Settings.OutlineColor),
                    TextSize = Settings.FontSize,
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = Settings.OutlineWidth
                };

                nameOutlinePaint = new SKPaint
                {
                    Color = ParseColor(Settings.OutlineColor),
                    TextSize = Settings.FontSize,
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Normal),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = Settings.OutlineWidth
                };
            }

            int currentY = startY;
            int rankOffset = GetRankOffset(Settings.RankFormat);

            foreach (var ranking in rankings)
            {
                string rankText = RankFormatHelper.FormatRank(ranking.Rank, Settings.RankFormat);

                if (Settings.HasTextOutline)
                {
                    canvas.DrawText(rankText, leftMargin, currentY, rankOutlinePaint);
                    canvas.DrawText(ranking.PlayerName, leftMargin + rankOffset, currentY, nameOutlinePaint);
                }

                canvas.DrawText(rankText, leftMargin, currentY, rankPaint);
                canvas.DrawText(ranking.PlayerName, leftMargin + rankOffset, currentY, namePaint);

                currentY += Settings.Spacing;
            }

            rankOutlinePaint?.Dispose();
            nameOutlinePaint?.Dispose();
        }

        private void DrawRankingsInColumns(SKCanvas canvas, List<RankingEntry> rankings, int startY, int totalWidth, int availableHeight)
        {
            using var rankPaint = new SKPaint
            {
                Color = ParseColor(Settings.RankNumberColor),
                TextSize = Settings.FontSize,
                IsAntialias = true,
                Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold)
            };

            using var namePaint = new SKPaint
            {
                Color = ParseColor(Settings.PlayerTextColor),
                TextSize = Settings.FontSize,
                IsAntialias = true,
                Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Normal)
            };

            SKPaint rankOutlinePaint = null;
            SKPaint nameOutlinePaint = null;

            if (Settings.HasTextOutline)
            {
                rankOutlinePaint = new SKPaint
                {
                    Color = ParseColor(Settings.OutlineColor),
                    TextSize = Settings.FontSize,
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = Settings.OutlineWidth
                };

                nameOutlinePaint = new SKPaint
                {
                    Color = ParseColor(Settings.OutlineColor),
                    TextSize = Settings.FontSize,
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Normal),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = Settings.OutlineWidth
                };
            }

            int rankOffset = GetRankOffset(Settings.RankFormat);

            int itemsPerColumn = (int)Math.Ceiling(availableHeight / (float)Settings.Spacing);
            int columnWidth = totalWidth / 2;

            int leftMargin1 = Settings.PaddingLeft + 20;
            int leftMargin2 = columnWidth + Settings.PaddingLeft + 20;

            for (int i = 0; i < rankings.Count; i++)
            {
                var ranking = rankings[i];
                bool isLeftColumn = i < itemsPerColumn;

                int leftMargin = isLeftColumn ? leftMargin1 : leftMargin2;
                int yIndex = isLeftColumn ? i : i - itemsPerColumn;
                int currentY = startY + (yIndex * Settings.Spacing);

                string rankText = RankFormatHelper.FormatRank(ranking.Rank, Settings.RankFormat);

                if (Settings.HasTextOutline)
                {
                    canvas.DrawText(rankText, leftMargin, currentY, rankOutlinePaint);
                    canvas.DrawText(ranking.PlayerName, leftMargin + rankOffset, currentY, nameOutlinePaint);
                }

                canvas.DrawText(rankText, leftMargin, currentY, rankPaint);
                canvas.DrawText(ranking.PlayerName, leftMargin + rankOffset, currentY, namePaint);
            }

            rankOutlinePaint?.Dispose();
            nameOutlinePaint?.Dispose();
        }

        private int GetRankOffset(RankFormat format)
        {
            return format switch
            {
                RankFormat.NumberDot => 60,
                RankFormat.HashDash => 90,
                RankFormat.PlaceColon => 150,
                _ => 60
            };
        }

        private void MarkDirty()
        {
            _isDirty = true;
            _saveTimer.Stop();
            _saveTimer.Start();
        }

        private void LoadData()
        {
            if (File.Exists(_dataPath))
            {
                try
                {
                    var json = File.ReadAllText(_dataPath);
                    var data = JsonSerializer.Deserialize<RankingsData>(json);

                    foreach (var entry in data.Rankings)
                    {
                        Rankings.Add(entry);
                    }
                }
                catch { }
            }

            // Ensure we have 20 entries
            while (Rankings.Count < 50)
            {
                Rankings.Add(new RankingEntry { Rank = Rankings.Count + 1, PlayerName = "" });
            }
        }

        private void SaveData()
        {
            if (!_isDirty) return;

            try
            {
                var data = new RankingsData { Rankings = Rankings.ToList() };
                var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_dataPath, json);
                _isDirty = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving data: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadSettings()
        {
            if (File.Exists(_settingsPath))
            {
                try
                {
                    var json = File.ReadAllText(_settingsPath);
                    Settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
                catch { }
            }
        }

        private void SaveSettings()
        {
            try
            {
                _settingsDirty = true; // Add this line at the top

                var json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsPath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private string GetPresetsFolder()
        {
            var appFolder = AppDomain.CurrentDomain.BaseDirectory;
            var presetsFolder = Path.Combine(appFolder, "SettingsPresets");
            Directory.CreateDirectory(presetsFolder);
            return presetsFolder;
        }

        private List<string> LoadPresetsList()
        {
            var presetsFolder = GetPresetsFolder();
            var files = Directory.GetFiles(presetsFolder, "*.json");
            return files.Select(f => Path.GetFileNameWithoutExtension(f)).ToList();
        }

        private void RebuildSettingsPanel()
        {
            // Use the class field _mainGrid if you have it, otherwise find it via DockPanel
            var mainGrid = _mainGrid;

            if (mainGrid == null && Content is DockPanel dockPanel)
            {
                mainGrid = dockPanel.Children.OfType<Grid>().FirstOrDefault();
            }

            if (mainGrid != null)
            {
                // 1. Find the specific element sitting in Column 2 (The Settings Panel)
                UIElement oldSettingsPanel = null;
                foreach (UIElement child in mainGrid.Children)
                {
                    if (Grid.GetColumn(child) == 2)
                    {
                        oldSettingsPanel = child;
                        break;
                    }
                }

                // 2. Remove it safely
                if (oldSettingsPanel != null)
                {
                    mainGrid.Children.Remove(oldSettingsPanel);
                }

                // 3. Create the new panel
                var newSettingsPanel = CreateSettingsPanel();

                // 4. Put it in COLUMN 2 (The Settings Column)
                Grid.SetColumn(newSettingsPanel, 2);

                // 5. Add it back to the grid
                mainGrid.Children.Add(newSettingsPanel);
            }
        }

        private void LoadPresetFromFile(string presetName)
        {
            try
            {
                var presetsFolder = GetPresetsFolder();
                var filePath = Path.Combine(presetsFolder, presetName + ".json");

                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"Preset file not found: {presetName}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var json = File.ReadAllText(filePath);
                var loadedSettings = JsonSerializer.Deserialize<AppSettings>(json);

                if (loadedSettings != null)
                {
                    // Copy all properties using reflection
                    foreach (var prop in typeof(AppSettings).GetProperties())
                    {
                        if (prop.CanWrite)
                        {
                            var value = prop.GetValue(loadedSettings);
                            prop.SetValue(Settings, value);
                        }
                    }

                    _currentPresetName = presetName;
                    _settingsDirty = false;
                    RebuildSettingsPanel();

                    MessageBox.Show($"Loaded preset: {presetName}", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading preset: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SavePresetDialog()
        {
            var presetsFolder = GetPresetsFolder();

            var dialog = new SaveFileDialog
            {
                Filter = "JSON Files|*.json",
                Title = "Save Settings Preset",
                InitialDirectory = presetsFolder,
                FileName = string.IsNullOrEmpty(_currentPresetName) ? "MyPreset" : _currentPresetName
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(dialog.FileName, json);

                    _currentPresetName = Path.GetFileNameWithoutExtension(dialog.FileName);
                    _settingsDirty = false;

                    // Refresh preset dropdown
                    RefreshPresetDropdown();

                    MessageBox.Show($"Preset saved: {_currentPresetName}", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving preset: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void RefreshPresetDropdown()
        {
            if (_presetComboBox == null) return;

            _presetComboBox.Items.Clear();
            // _presetComboBox.Items.Add("Current Settings");

            var presets = LoadPresetsList();
            foreach (var preset in presets)
            {
                _presetComboBox.Items.Add(preset);
            }

            // Select current preset if one is loaded
            if (!string.IsNullOrEmpty(_currentPresetName) && presets.Contains(_currentPresetName))
            {
                _presetComboBox.SelectedItem = _currentPresetName;
            }
            else
            {
                _presetComboBox.SelectedIndex = 0; // "Current Settings"
            }
        }


    }

}