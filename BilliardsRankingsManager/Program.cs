using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using SkiaSharp;

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
        public static Brush UIAccent = new SolidColorBrush(Color.FromRgb(27, 94, 32));
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
        

        private void InitializeComponent()
        {
            Title = "Billiards Top 20 Rankings Manager";
            Width = 700;
            Height = 700;
            Background = StyleConfig.UIBackground;

            var mainGrid = new Grid();
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.5, GridUnitType.Star) });

            // Left Panel - Rankings
            var leftPanel = CreateRankingsPanel();
            Grid.SetColumn(leftPanel, 0);
            mainGrid.Children.Add(leftPanel);

            // Right Panel - Settings and Export
            var rightPanel = CreateSettingsPanel();
            Grid.SetColumn(rightPanel, 1);
            mainGrid.Children.Add(rightPanel);

            Content = mainGrid;
        }

        private ScrollViewer CreateRankingsPanel()
        {
            var stack = new StackPanel { Margin = new Thickness(15) }; // Reduced from 20

            var title = new TextBlock
            {
                Text = "Player Rankings",
                FontSize = 18, // Reduced from 24
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 10), // Reduced from 20
                Foreground = StyleConfig.UIAccent
            };
            stack.Children.Add(title);

            // Create 20 ranking entries
            for (int i = 0; i < 20; i++)
            {
                var entryGrid = new Grid { Margin = new Thickness(0, 2, 0, 2) }; // Reduced from 5
                entryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                entryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var rankLabel = new TextBlock
                {
                    Text = $"{i + 1}.",
                    Width = 30, // Reduced from 40
                    FontSize = 12, // Reduced from 16
                    FontWeight = FontWeights.Bold,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(rankLabel, 0);

                var nameBox = new TextBox
                {
                    FontSize = 13, // Reduced from 16
                    Padding = new Thickness(3), // Reduced from 8
                    Margin = new Thickness(5, 0, 0, 0)
                };
                nameBox.SetBinding(TextBox.TextProperty, new System.Windows.Data.Binding("PlayerName")
                {
                    Source = Rankings[i],
                    Mode = System.Windows.Data.BindingMode.TwoWay,
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                });
                Grid.SetColumn(nameBox, 1);

                entryGrid.Children.Add(rankLabel);
                entryGrid.Children.Add(nameBox);
                stack.Children.Add(entryGrid);
            }

            return new ScrollViewer
            {
                Content = stack,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
        }

        private ScrollViewer CreateSettingsPanel()
        {
            var stack = new StackPanel { Margin = new Thickness(15) }; // Reduced from 20

            // Settings Title
            var title = new TextBlock
            {
                Text = "Settings & Export",
                FontSize = 18, // Reduced from 20
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 15), // Reduced from 20
                Foreground = StyleConfig.UIAccent
            };
            stack.Children.Add(title);

            // Export Settings Section
            stack.Children.Add(CreateSectionHeader("Export Settings"));

            stack.Children.Add(CreateNumberSetting("Width (px):", Settings.ExportWidth,
                v => { Settings.ExportWidth = v; SaveSettings(); }));
            stack.Children.Add(CreateNumberSetting("Height (px):", Settings.ExportHeight,
                v => { Settings.ExportHeight = v; SaveSettings(); }));

            // Logo Settings Section
            stack.Children.Add(CreateSectionHeader("Logo"));

            var logoBtn = CreateButton("Browse for Logo", BrowseLogo_Click);
            logoBtn.Margin = new Thickness(0, 0, 0, 5);
            stack.Children.Add(logoBtn);

            var logoStatus = new TextBlock
            {
                Text = string.IsNullOrEmpty(Settings.LogoPath) ? "No logo selected" : Path.GetFileName(Settings.LogoPath),
                Margin = new Thickness(0, 0, 0, 10),
                FontSize = 11,
                FontStyle = FontStyles.Italic,
                TextWrapping = TextWrapping.Wrap
            };
            logoStatus.Name = "LogoStatus";
            stack.Children.Add(logoStatus);

            stack.Children.Add(CreateSliderSetting("Logo Size:", Settings.LogoHeight, 50, 250,
                v => { Settings.LogoHeight = v; SaveSettings(); }));

            stack.Children.Add(CreateSliderSetting("Logo-Title Spacing:", Settings.LogoTitleSpacing, 20, 100,
                v => { Settings.LogoTitleSpacing = v; SaveSettings(); }));

            // Title Settings Section
            stack.Children.Add(CreateSectionHeader("Title"));

            var titlePanel = new StackPanel { Margin = new Thickness(0, 5, 0, 10) };
            titlePanel.Children.Add(new TextBlock { Text = "Title Text:", FontWeight = FontWeights.SemiBold, FontSize = 12 });
            var titleBox = new TextBox
            {
                Text = Settings.TitleText,
                Margin = new Thickness(0, 5, 0, 0),
                Padding = new Thickness(5)
            };
            titleBox.TextChanged += (s, e) =>
            {
                Settings.TitleText = titleBox.Text;
                SaveSettings();
            };
            titlePanel.Children.Add(titleBox);
            stack.Children.Add(titlePanel);

 
            // Layout Settings Section
            stack.Children.Add(CreateSectionHeader("Layout"));

            // Colors Section - Collapsible
            var colorsExpander = new Expander
            {
                Header = "Colors",
                IsExpanded = false,
                Margin = new Thickness(0, 15, 0, 10),
                FontWeight = FontWeights.Regular,
                FontSize = 14
            };

            var colorsStack = new StackPanel();
            colorsStack.Children.Add(CreateColorSetting("Background:", Settings.BackgroundColor,
                c => { Settings.BackgroundColor = c; SaveSettings(); }));
            colorsStack.Children.Add(CreateColorSetting("Border:", Settings.BorderColor,
                c => { Settings.BorderColor = c; SaveSettings(); }));
            colorsStack.Children.Add(CreateColorSetting("Title Text:", Settings.TitleTextColor,
                c => { Settings.TitleTextColor = c; SaveSettings(); }));
            colorsStack.Children.Add(CreateColorSetting("Player Text:", Settings.PlayerTextColor,
                c => { Settings.PlayerTextColor = c; SaveSettings(); }));
            colorsStack.Children.Add(CreateColorSetting("Rank Numbers:", Settings.RankNumberColor,
                c => { Settings.RankNumberColor = c; SaveSettings(); }));

            colorsExpander.Content = colorsStack;
            stack.Children.Add(colorsExpander);

            // Margin/padding Section - Collapsible
            var paddingExpander = new Expander
            {
                Header = "Margins",
                IsExpanded = false,
                Margin = new Thickness(0, 15, 0, 10),
                FontWeight = FontWeights.Regular,
                FontSize = 14
            };

            var marginStack = new StackPanel();
            marginStack.Children.Add(CreateNumberSetting("Margin Top:", Settings.PaddingTop,
                v => { Settings.PaddingTop = v; SaveSettings(); }));
            marginStack.Children.Add(CreateNumberSetting("Margin Bottom:", Settings.PaddingBottom,
                v => { Settings.PaddingBottom = v; SaveSettings(); }));
            marginStack.Children.Add(CreateNumberSetting("Margin Left:", Settings.PaddingLeft,
                v => { Settings.PaddingLeft = v; SaveSettings(); }));
            marginStack.Children.Add(CreateNumberSetting("Margin Right:", Settings.PaddingRight,
                v => { Settings.PaddingRight = v; SaveSettings(); }));

            paddingExpander.Content = marginStack;
            stack.Children.Add(paddingExpander);

            stack.Children.Add(CreateNumberSetting("Font Size:", Settings.FontSize,
                v => { Settings.FontSize = v; SaveSettings(); }));
            stack.Children.Add(CreateNumberSetting("Spacing:", Settings.Spacing,
                v => { Settings.Spacing = v; SaveSettings(); }));



            // Text Outline Section
            stack.Children.Add(CreateSectionHeader("Text Outline"));

            var outlineCheckPanel = new StackPanel { Margin = new Thickness(0, 5, 0, 10) };
            var outlineCheck = new CheckBox
            {
                Content = "Enable Text Outline",
                IsChecked = Settings.HasTextOutline
            };
            outlineCheck.Checked += (s, e) => { Settings.HasTextOutline = true; SaveSettings(); };
            outlineCheck.Unchecked += (s, e) => { Settings.HasTextOutline = false; SaveSettings(); };
            outlineCheckPanel.Children.Add(outlineCheck);
            stack.Children.Add(outlineCheckPanel);

            stack.Children.Add(CreateColorSetting("Outline Color:", Settings.OutlineColor,
                c => { Settings.OutlineColor = c; SaveSettings(); }));
            stack.Children.Add(CreateSliderSetting("Outline Width:", Settings.OutlineWidth, 1, 10,
                v => { Settings.OutlineWidth = v; SaveSettings(); }));


            // Font Family
            var fontPanel = new StackPanel { Margin = new Thickness(0, 10, 0, 10) };
            fontPanel.Children.Add(new TextBlock { Text = "Font Family:", FontWeight = FontWeights.SemiBold, FontSize = 12 });
            var fontBox = new TextBox
            {
                Text = Settings.FontFamily,
                Margin = new Thickness(0, 5, 0, 0),
                Padding = new Thickness(5)
            };
            fontBox.TextChanged += (s, e) =>
            {
                Settings.FontFamily = fontBox.Text;
                SaveSettings();
            };
            fontPanel.Children.Add(fontBox);
            stack.Children.Add(fontPanel);

            // Rank Format Dropdown
            var rankFormatPanel = new StackPanel { Margin = new Thickness(0, 10, 0, 10) };
            rankFormatPanel.Children.Add(new TextBlock { Text = "Rank Format:", FontWeight = FontWeights.SemiBold, FontSize = 12 });

            var rankFormatCombo = new ComboBox
            {
                Margin = new Thickness(0, 5, 0, 0),
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
            stack.Children.Add(rankFormatPanel);

            // Export Location Section
            stack.Children.Add(CreateSectionHeader("Export"));

            var locationBtn = CreateButton("Set Export Location", SetExportLocation_Click);
            locationBtn.Margin = new Thickness(0, 0, 0, 5);
            stack.Children.Add(locationBtn);

            var locationStatus = new TextBlock
            {
                Text = string.IsNullOrEmpty(Settings.ExportPath) ? "No location set" : Settings.ExportPath,
                Margin = new Thickness(0, 0, 0, 15),
                FontSize = 11,
                FontStyle = FontStyles.Italic,
                TextWrapping = TextWrapping.Wrap
            };
            locationStatus.Name = "LocationStatus";
            stack.Children.Add(locationStatus);

            // Export Buttons
            var previewBtn = CreateButton("Preview", (s, e) => _ = ShowPreview(), isPrimary: false);
            previewBtn.Name = "PreviewButton";
            stack.Children.Add(previewBtn);

            var pngBtn = CreateButton("Generate PNG", (s, e) => _ = ExportToPNG(), isPrimary: true);
            pngBtn.Margin = new Thickness(0, 10, 0, 0);
            stack.Children.Add(pngBtn);

            return new ScrollViewer
            {
                Content = stack,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
        }


        private TextBlock CreateSectionHeader(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 14, // Reduced from 16
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 12, 0, 8) // Reduced from 15,10
            };
        }

        private StackPanel CreateNumberSetting(string label, int defaultValue, Action<int> onChange)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 5, 0, 5) };
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

        private StackPanel CreateSliderSetting(string label, int defaultValue, int min, int max, Action<int> onChange)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 5, 0, 5) };

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
                TickFrequency = 1,
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

        private StackPanel CreateColorSetting(string label, string currentColor, Action<string> onChange)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 5, 0, 5), Orientation = Orientation.Horizontal };

            var labelText = new TextBlock
            {
                Text = label,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Width = 120
            };
            panel.Children.Add(labelText);

            var colorButton = new Button
            {
                Width = 80,
                Height = 25,
                Content = "Pick Color",
                Margin = new Thickness(5, 0, 0, 0)
            };

            var previewBox = new Border
            {
                Width = 40,
                Height = 25,
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Margin = new Thickness(5, 0, 0, 0)
            };

            Action updatePreview = () =>
            {
                try
                {
                    if (currentColor.Equals("transparent", StringComparison.OrdinalIgnoreCase))
                    {
                        previewBox.Background = new SolidColorBrush(Colors.Transparent);
                    }
                    else
                    {
                        var color = (Color)ColorConverter.ConvertFromString(currentColor);
                        previewBox.Background = new SolidColorBrush(color);
                    }
                }
                catch
                {
                    previewBox.Background = Brushes.White;
                }
            };

            updatePreview();

            colorButton.Click += (s, e) =>
            {
                var dialog = new System.Windows.Forms.ColorDialog
                {
                    FullOpen = true,
                    AnyColor = true
                };

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var wpfColor = Color.FromArgb(dialog.Color.A, dialog.Color.R, dialog.Color.G, dialog.Color.B);
                    currentColor = wpfColor.ToString();
                    onChange(currentColor);
                    updatePreview();
                }
            };

            var transparentBtn = new Button
            {
                Content = "Transparent",
                Width = 80,
                Height = 25,
                Margin = new Thickness(5, 0, 0, 0)
            };

            transparentBtn.Click += (s, e) =>
            {
                currentColor = "Transparent";
                onChange(currentColor);
                updatePreview();
            };

            panel.Children.Add(colorButton);
            panel.Children.Add(transparentBtn);
            panel.Children.Add(previewBox);

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
                button.Background = StyleConfig.UIAccent;
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
                TextSize = StyleConfig.DefaultTitleFontSize,
                IsAntialias = true,
                Typeface = SKTypeface.FromFamilyName(Settings.FontFamily, SKFontStyle.Bold),
                TextAlign = SKTextAlign.Center
            };

            if (Settings.HasTextOutline)
            {
                using var outlinePaint = new SKPaint
                {
                    Color = ParseColor(Settings.OutlineColor),
                    TextSize = StyleConfig.DefaultTitleFontSize,
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
            var activeRankings = Rankings.Where(r => !string.IsNullOrWhiteSpace(r.PlayerName))
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
                    Width = 900,
                    Height = 700,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
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
            while (Rankings.Count < 20)
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
                var json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsPath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            SaveData();

            // Cleanup preview resources
            if (_previewUpdateTimer != null)
            {
                _previewUpdateTimer.Stop();
                _previewUpdateTimer.Dispose();
            }

            if (_previewWindow != null && _previewWindow.IsVisible)
            {
                _previewWindow.Close();
            }

            base.OnClosing(e);
        }
    }


    public class App : Application
    {
        [STAThread]
        public static void Main()
        {
            var app = new App();
            app.Run(new MainWindow());
        }
    }
}