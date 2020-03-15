using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Xml;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using Microsoft.WindowsAPICodePack.Dialogs;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using System.Security.Cryptography;
using System.Windows.Documents;
using System.Collections.Generic;

namespace Excel2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private Grid mMainGrid;

        //private TextEditor mDotTemplate_TextBox;

        private readonly DataManages mDataManages;
        private readonly ObservableCollection<ListViewItemData> ListViweItemData;

        public string FileName { get; set; }

        /// <summary>
        /// 表头数
        /// </summary>
        private int HeadNum { get; set; }

        /// <summary>
        /// 是否支持多页签
        /// </summary>
        private bool MultiSheet { get; set; }

        /// <summary>
        /// CSRadioBtn is Checked 
        /// </summary>
        private bool CSRadioBtnChecked { get; set; }

        /// <summary>
        /// CSRadioBtn is Checked 
        /// </summary>
        private bool TSRadioBtnChecked { get; set; }

        private bool JsonRadioBtnChecked { get; set; }
        private bool TemplateRadioBtnChecked { get; set; }

        /// <summary>
        /// 导出模板文件类型
        /// </summary>
        private TemplateType Type { get; set; }

        private string JsonData { get; set; }

        private string TemplateData { get; set; }

        private readonly BackgroundWorker mBgworker;
        private readonly DoWorkEventHandler mDoWorkEventHandler;
        private readonly ProgressChangedEventHandler mProgressChangedEventHandler;
        private readonly CommonOpenFileDialog mFolderDialog;

        private readonly BackgroundWorker mBgShowFileList;
        private readonly DoWorkEventHandler mDoShowFileHandler;

        private readonly IHighlightingDefinition JsonHighlighting;
        private readonly IHighlightingDefinition CSHighlighting;
        private readonly IHighlightingDefinition TSHighlighting;

        private TextBoxData ExcelPath;
        private TextBoxData JsonPath;
        private TextBoxData TemplatePath;
        private TextBoxData SheetSign;
        private TextBoxData EncryptionKey;
        private TextBoxData EncryptionIV;

        public MainWindow()
        {
            InitializeComponent();
            this.Closing += Window_Closing;
            this.Loaded += Window_Loaded;

            mDataManages = new DataManages();
            mFolderDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true//设置为选择文件夹
            };
            mBgworker = new BackgroundWorker();
            mDoWorkEventHandler = new DoWorkEventHandler(DoWork);
            mProgressChangedEventHandler = new ProgressChangedEventHandler(BgworkChange);

            mBgShowFileList = new BackgroundWorker();
            mDoShowFileHandler = new DoWorkEventHandler(DoShowFileList);
            mBgShowFileList.DoWork += mDoShowFileHandler;

            ListViweItemData = new ObservableCollection<ListViewItemData>();

            ExcelPath = new TextBoxData();
            JsonPath = new TextBoxData();
            TemplatePath = new TextBoxData();
            SheetSign = new TextBoxData();
            EncryptionKey = new TextBoxData();
            EncryptionIV = new TextBoxData();

            ExcelPath.Text = Properties.Settings.Default.ExcelPath;
            JsonPath.Text = Properties.Settings.Default.JsonPath;
            TemplatePath.Text = Properties.Settings.Default.TemplatePath;
            HeadNum = Properties.Settings.Default.HeadNum;
            MultiSheet = Properties.Settings.Default.MultiSheet;
            CSRadioBtnChecked = Properties.Settings.Default.CSRadioBtnChecked;
            TSRadioBtnChecked = Properties.Settings.Default.TSRadioBtnChecked;

            Type = TemplateType.MIN;
            if (CSRadioBtnChecked) Type = TemplateType.CS;
            if (TSRadioBtnChecked) Type = TemplateType.TS;
            if (!CSRadioBtnChecked && !TSRadioBtnChecked) Type = TemplateType.CS;


            StreamReader stream = new StreamReader(Application.GetResourceStream(new Uri("Resources/JsonDark.xshd", UriKind.Relative)).Stream, Encoding.UTF8);
            using (XmlTextReader reader = new XmlTextReader(stream))
            {
                JsonHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                reader.Close();
                stream.Close();
            }
            stream = new StreamReader(Application.GetResourceStream(new Uri("Resources/CSDark.xshd", UriKind.Relative)).Stream, Encoding.UTF8);
            using (XmlTextReader reader = new XmlTextReader(stream))
            {
                CSHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                reader.Close();
                stream.Close();
            }
            stream = new StreamReader(Application.GetResourceStream(new Uri("Resources/TSDark.xshd", UriKind.Relative)).Stream, Encoding.UTF8);
            using (XmlTextReader reader = new XmlTextReader(stream))
            {
                TSHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                reader.Close();
                stream.Close();
            }
            //this.Width = this.MinWidth;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //this.Width = this.MinWidth;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            mBgShowFileList.DoWork -= mDoWorkEventHandler;
            Properties.Settings.Default.Save();
        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            mMainGrid = sender as Grid;

            ExcelListView.ItemsSource = ListViweItemData;

            ExcelPath_TextBox.DataContext = ExcelPath;
            JsonPath_TextBox.DataContext = JsonPath;
            DotTemplateFilePath_TextBox.DataContext = TemplatePath;
            Signsheet_Textbox.DataContext = SheetSign;
            EncryptionKey_Textbox.DataContext = EncryptionKey;
            EncryptionIV_Textbox.DataContext = EncryptionIV;

            if (BeginBtn != null)
            {
                BeginBtn.Content = "Begin";
                BeginBtn.Click += (s, ee) => Button_ClickAsync(s, ee);
            }

            FilterNum_ComboBox.SelectedIndex = HeadNum - 1;

            ProgressBar.Value = 0;

            Mutilsheet_Checkbox.IsChecked = MultiSheet;
            Signsheet_Textbox.Visibility = MultiSheet ? Visibility.Visible : Visibility.Hidden;

            EncryptionKey.Text = "";
            EncryptionIV.Text = "";

            Dotcs_RadioBtn.IsChecked = CSRadioBtnChecked;
            Dotts_RadioBtn.IsChecked = TSRadioBtnChecked;

            if (!CSRadioBtnChecked && !TSRadioBtnChecked) Dotcs_RadioBtn.IsChecked = true;

            TextView.SyntaxHighlighting = JsonHighlighting;

            SetColor(Properties.Settings.Default.Color, Properties.Settings.Default.Theme);

            SetEncryptionUI(false);
            MutilSheet_Label.IsEnabled = MultiSheet;


            List<ThemesListBoxItem> list = new List<ThemesListBoxItem>();
            ThemesListbox.ItemsSource = list;

            foreach (var item in Enum.GetValues(typeof(Themes)))
            {
                list.Add(new ThemesListBoxItem(item.ToString()));
            };
        }

        private void Button_ClickAsync(object sender, RoutedEventArgs e)
        {
            bool isExistExclePath = true;
            bool isExistJsonPath = true;
            bool isExistTemplatePath = true;


            if (!Directory.Exists(ExcelPath.Text))
                isExistExclePath = false;
            if (!Directory.Exists(JsonPath.Text))
                isExistJsonPath = false;
            if (!Directory.Exists(TemplatePath.Text))
                isExistTemplatePath = false;

            if (!isExistExclePath || !isExistJsonPath)
            {
                if (!isExistExclePath)
                    _ = MessageBox.Show("ExcelPath does not exist!", "Directory not exist!");
                else if (!isExistJsonPath)
                    _ = MessageBox.Show("JsonPath does not exist!", "Directory not exist");
                SetUIStates(true);
            }
            else
            {
                ProgressBar.Maximum = mDataManages.FilesCount(HeadNum, isExistTemplatePath);
                ProgressBar.Value = 0;
                mBgworker.WorkerReportsProgress = true;
                mBgworker.DoWork += mDoWorkEventHandler;
                mBgworker.ProgressChanged += mProgressChangedEventHandler;
                mBgworker.RunWorkerAsync();

                SetUIStates(false);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonFileDialogResult result = mFolderDialog.ShowDialog();
            if (result == CommonFileDialogResult.Cancel)
            {
                return;
            }
            //string s = mFolderDialog.FileName;
            //Trace.WriteLine(s);
            if (((Button)sender).Name.Equals("excelpath_btn"))
                ExcelPath.Text = mFolderDialog.FileName;// FolderDialog.SelectedPath.Trim();

            if (((Button)sender).Name.Equals("jsonpath_btn"))
                JsonPath_TextBox.Text = mFolderDialog.FileName; // FolderDialog.SelectedPath.Trim();

            if (((Button)sender).Name.Equals("templatefilepath_btn"))
                DotTemplateFilePath_TextBox.Text = mFolderDialog.FileName; // FolderDialog.SelectedPath.Trim();
        }

        private void ToggleButton_Checked(object sender, RoutedEventArgs e)
        {
            ToggleButton cb = sender as ToggleButton;
            switch (cb.Name)
            {
                case "Encryption_ToggleBtn":
                    bool isChecked = (bool)cb.IsChecked;
                    mDataManages.CanEncryption = isChecked;
                    SetEncryptionUI(isChecked);
                    break;
                case "Mutilsheet_Checkbox":
                    MultiSheet = (bool)cb.IsChecked;
                    Properties.Settings.Default.MultiSheet = MultiSheet;
                    Signsheet_Textbox.Visibility = MultiSheet ? Visibility.Visible : Visibility.Hidden;
                    MutilSheet_Label.IsEnabled = MultiSheet;
                    ShowFileList();
                    break;
            }
        }

        private void SetEncryptionUI(bool _isChecked)
        {
            List<ComboxEncryptionMode> list = new List<ComboxEncryptionMode>();
            int modeID = 0;
            foreach (var item in Enum.GetValues(typeof(EncryptionMode)))
            {
                modeID++;
                list.Add(new ComboxEncryptionMode(item.ToString(), modeID.ToString()));
            }

            List<ComboxEncryptionPadding> listPadding = new List<ComboxEncryptionPadding>();
            int paddingID = 0;
            foreach (var item in Enum.GetValues(typeof(EncryptionPadding)))
            {
                paddingID++;
                listPadding.Add(new ComboxEncryptionPadding(item.ToString(), paddingID.ToString()));
            }

            Encryption_Mode_ComboBox.ItemsSource = list;
            Encryption_Padding_ComboBox.ItemsSource = listPadding;

            Encryption_ToggleBtn.IsChecked = _isChecked;
            Encryption_Label.IsEnabled = _isChecked;
            Encryption_Mode_Label.IsEnabled = _isChecked;
            Encryption_Padding_Label.IsEnabled = _isChecked;
            Encryption_Mode_ComboBox.IsEnabled = _isChecked;
            Encryption_Padding_ComboBox.IsEnabled = _isChecked;
            EncryptionKey_Textbox.IsEnabled = _isChecked;
            EncryptionIV_Textbox.IsEnabled = _isChecked;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = sender as ComboBox;

            switch (cb.Name)
            {
                case "FilterNum_ComboBox":
                    HeadNum = cb.SelectedIndex + 1;
                    if (Properties.Settings.Default.HeadNum != HeadNum)
                        ShowFileList();

                    if (HeadNum > 1)
                    {
                        Dotcs_RadioBtn.Visibility = Visibility.Visible;
                        Dotts_RadioBtn.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        Dotcs_RadioBtn.Visibility = Visibility.Hidden;
                        Dotts_RadioBtn.Visibility = Visibility.Hidden;
                    }

                    Properties.Settings.Default.HeadNum = HeadNum;
                    break;
                case "Encryption_Mode_ComboBox":
                    mDataManages.Mode = (CipherMode)Enum.Parse(typeof(CipherMode), cb.SelectedIndex + 1 + "");
                    break;
                case "Encryption_Padding_ComboBox":
                    mDataManages.Padding = (PaddingMode)Enum.Parse(typeof(PaddingMode), cb.SelectedIndex + 1 + "");
                    break;
            }
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            UpdateText(tb, tb.Text);
        }

        private void Textbox_DragEnter(object sender, DragEventArgs e)
        {
            //TextBox tb = sender as TextBox;
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void Textbox_DragDrop(object sender, DragEventArgs e)
        {
            TextBox tb = sender as TextBox;
            string fileName = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            tb.Text = fileName;
            UpdateText(tb, tb.Text);
        }

        private void Textbox_LostFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            //TextBox cb = sender as TextBox;
            //if (cb.Name.Equals("exclepath_textbox")) ShowFileList();
            //else if (cb.Name.Equals("jsonpath_textbox")) { }
            //else if (cb.Name.Equals("templatefilepath_textbox")) { }
            //else if (cb.Name.Equals("signsheet_textbox")) ShowFileList();
        }

        private void ListView_MouseClick(object sender, RoutedEventArgs e)
        {
            if (!(ExcelListView.SelectedItem is ListViewItemData obj)) return;
            //this.Width = 1220;
            string name = obj.FileInfo.Name.Split('.')[0];
            string json = "";
            string template = "";
            if (!MultiSheet)
            {
                if (mDataManages.JsonData.ContainsKey(name))
                    json = mDataManages.JsonData[name];
                if (mDataManages.TemplateData.ContainsKey(name))
                    template = mDataManages.TemplateData[name];
            }
            else
            {
                json = "";
                foreach (var item in mDataManages.JsonData)
                {
                    if (item.Key.Contains(name) && item.Value != null)
                    {
                        json += "\n======> " + item.Key + " <======\n";
                        json += item.Value;
                    }
                }

                template = "";
                foreach (var item in mDataManages.TemplateData)
                {
                    if (item.Key.Contains(name) && item.Value != null)
                    {
                        template += "\n======> " + item.Key + " <======\n";
                        template += item.Value;
                    }
                }
            }

            JsonData = json;
            TemplateData = template;
            if (JsonRadioBtnChecked) TextView.Text = JsonData;
            if (TemplateRadioBtnChecked) TextView.Text = TemplateData;
            //mDotTemplate_TextBox.Text = template;
        }

        private void MenuItem_RightClick(object sender, RoutedEventArgs e)
        {
            if (!(ExcelListView.SelectedItem is ListViewItemData obj)) return;
            //this.Width = 1220;
            string name = obj.FileInfo.Name.Split('.')[0];
            if (!MultiSheet)
            {
                mDataManages.SaveFile(JsonPath.Text, TemplatePath.Text, HeadNum, Type, name, null);
            }
            else
            {
                foreach (var item in mDataManages.JsonData)
                {
                    if (item.Key.Contains(name) && item.Value != null)
                    {
                        mDataManages.SaveFile(JsonPath.Text, TemplatePath.Text, HeadNum, Type, item.Key, null);
                    }
                }
            }
        }

        private void Radiobtn_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton rb = sender as RadioButton;
            switch (rb.Name)
            {
                case "Dotts_RadioBtn":
                    Type = TemplateType.TS;
                    CSRadioBtnChecked = false;
                    TSRadioBtnChecked = true;
                    break;
                case "Dotcs_RadioBtn":
                    Type = TemplateType.CS;
                    CSRadioBtnChecked = true;
                    TSRadioBtnChecked = false;
                    break;
                case "JsonView_RadioBtn":
                    TemplateRadioBtnChecked = false;
                    JsonRadioBtnChecked = true;
                    if (TextView != null)
                    {
                        TextView.SyntaxHighlighting = JsonHighlighting;
                        TextView.Text = JsonData;
                    }
                    return;
                case "TemplateView_RadioBtn":
                    TemplateRadioBtnChecked = true;
                    JsonRadioBtnChecked = false;
                    if (TextView != null)
                    {
                        if (CSRadioBtnChecked)
                            TextView.SyntaxHighlighting = CSHighlighting;
                        else if (TSRadioBtnChecked)
                            TextView.SyntaxHighlighting = TSHighlighting;
                        TextView.Text = TemplateData;
                    }
                    return;
            }
            Properties.Settings.Default.CSRadioBtnChecked = CSRadioBtnChecked;
            Properties.Settings.Default.TSRadioBtnChecked = TSRadioBtnChecked;
            ShowFileList();
        }

        private void ListBoxItem_SelectionChanged(object sender, RoutedEventArgs e)
        {
            object ob = (sender as ListBox).SelectedItem;
            ThemesListBoxItem item = ob as ThemesListBoxItem;
            if (item.Name.Equals("Light") || item.Name.Equals("Dark"))
                Properties.Settings.Default.Theme = item.Name;
            else
                Properties.Settings.Default.Color = item.Name;
            SetColor(Properties.Settings.Default.Color, Properties.Settings.Default.Theme);
        }

        // =======================================

        private void BgworkChange(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value += e.ProgressPercentage;
            if (BeginBtn != null)
                BeginBtn.Content = $"{FileName} Finished";

            if (ProgressBar.Value >= ProgressBar.Maximum)
            {
                MessageBox.Show("Finished!");
                mBgworker.DoWork -= mDoWorkEventHandler;
                mBgworker.ProgressChanged -= mProgressChangedEventHandler;
                SetUIStates(true);
            }
        }

        private void DoWork(object sender, DoWorkEventArgs e)
        {
            if (Directory.Exists(JsonPath.Text))
            {
                mDataManages.SaveFiles(JsonPath.Text, TemplatePath.Text, HeadNum, Type, (d, v) =>
                 {
                     mBgworker.ReportProgress((int)d);
                     FileName = v;
                 });
            }

        }

        private void SetUIStates(bool _isEnable)
        {

            if (BeginBtn != null)
            {
                BeginBtn.Content = _isEnable ? "Begin" : "";
                BeginBtn.IsEnabled = _isEnable;
            }
            if (ExcelPathBtn != null)
                ExcelPathBtn.IsEnabled = _isEnable;

            if (JsonPathBtn != null)
                JsonPathBtn.IsEnabled = _isEnable;

            if (DotTemplateFilePathBtn != null)
                DotTemplateFilePathBtn.IsEnabled = _isEnable;

            if (ExcelPath_TextBox != null)
                ExcelPath_TextBox.IsEnabled = _isEnable;

            if (JsonPath_TextBox != null)
                JsonPath_TextBox.IsEnabled = _isEnable;

            if (DotTemplateFilePath_TextBox != null)
                DotTemplateFilePath_TextBox.IsEnabled = _isEnable;
        }

        private void DoShowFileList(object sender, DoWorkEventArgs e)
        {
            if (Directory.Exists(ExcelPath.Text) && ExcelListView != null)
            {
                DirectoryInfo TheFolder = new DirectoryInfo(ExcelPath.Text);
                foreach (var item in TheFolder.GetFiles())
                {
                    if (item.Extension.Equals(".xlsx") || item.Extension.Equals(".xls"))
                    {
                        mDataManages.ReadExcel(item);
                        mDataManages.ExportJson(item, HeadNum, MultiSheet, SheetSign.Text);
                        mDataManages.ExportTemplate(item, HeadNum, MultiSheet, Type, SheetSign.Text);
                    }
                }
            }
        }

        private void ShowFileList()
        {
            TextView?.Clear();
            ListViweItemData?.Clear();
            mDataManages?.ClearData();
            JsonRadioBtnChecked = true;
            TemplateRadioBtnChecked = false;
            JsonView_RadioBtn.IsChecked = JsonRadioBtnChecked;
            TemplateView_RadioBtn.IsChecked = TemplateRadioBtnChecked;
            //this.Width = this.MinWidth;

            if (Directory.Exists(ExcelPath.Text) && ExcelListView != null)
            {
                DirectoryInfo TheFolder = new DirectoryInfo(ExcelPath.Text);
                int idx = 0;
                foreach (var item in TheFolder.GetFiles())
                {
                    if (item.Extension.Equals(".xlsx") || item.Extension.Equals(".xls"))
                    {
                        ListViweItemData.Add(new ListViewItemData((idx++).ToString(), item));
                    }
                }
            }
            if (mBgShowFileList != null && !mBgShowFileList.IsBusy)
                mBgShowFileList.RunWorkerAsync();

            if (ProgressBar != null)
                ProgressBar.Value = 0;
        }

        private void SetColor(string _colorName, string _theme)
        {
            IBaseTheme theme = new MaterialDesignDarkTheme();
            MaterialDesignColor materialDesignColor = (MaterialDesignColor)Enum.Parse(typeof(MaterialDesignColor), _colorName);
            Color primaryColor = SwatchHelper.Lookup[materialDesignColor];
            Color accentColor = SwatchHelper.Lookup[materialDesignColor];

            switch (_theme)
            {
                case "Light":
                    theme = new MaterialDesignLightTheme();
                    TextView.Foreground = Brushes.Black;
                    break;
                case "Dark":
                    theme = new MaterialDesignDarkTheme();
                    TextView.Foreground = Brushes.White;
                    break;
            }

            ITheme themes = Theme.Create(theme, primaryColor, accentColor);
            Resources.SetTheme(themes);
        }

        private void UpdateText(TextBox _textBox, string _name)
        {
            switch (_textBox.Name)
            {
                case "ExcelPath_TextBox":
                    ExcelPath.Text = _name;
                    if (Directory.Exists(ExcelPath.Text) || string.IsNullOrEmpty(ExcelPath.Text))
                    {
                        Properties.Settings.Default.ExcelPath = ExcelPath.Text;
                        ShowFileList();
                    }
                    break;
                case "JsonPath_TextBox":
                    JsonPath.Text = _name;
                    if (Directory.Exists(JsonPath.Text) || string.IsNullOrEmpty(JsonPath.Text))
                        Properties.Settings.Default.JsonPath = JsonPath.Text;
                    break;
                case "DotTemplateFilePath_TextBox":
                    TemplatePath.Text = _name;
                    if (Directory.Exists(TemplatePath.Text) || string.IsNullOrEmpty(TemplatePath.Text))
                        Properties.Settings.Default.TemplatePath = TemplatePath.Text;
                    break;
                case "Signsheet_Textbox":
                    SheetSign.Text = _name;
                    ShowFileList();
                    break;
                case "EncryptionKey_Textbox":
                    EncryptionKey.Text = _name;
                    mDataManages.Key = _name;
                    break;
                case "EncryptionIV_Textbox":
                    EncryptionIV.Text = _name;
                    mDataManages.IV = _name;
                    break;
            }
        }

    }
}