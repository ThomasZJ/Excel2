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
        private Button mExcelPath_btn;
        private Button mJsonPath_btn;
        private Button mDotTemplateFilePath_btn;
        private Button mBegin_btn;

        private TextBox mExcelPath_TextBox;
        private TextBox mJsonPath_TextBox;
        private TextBox mDotTemplateFilePath_TextBox;
        private TextBox mSignsheet_Textbox;
        private TextBox mEncryptionKey_Textbox;
        private TextBox mEncryptionIV_Textbox;

        private RadioButton mDotcs_RadioBtn;
        private RadioButton mDotts_RadioBtn;
        private RadioButton mJsonView_RadioBtn;
        private RadioButton mTemplateView_RadioBtn;

        private ToggleButton mMutilsheet_Checkbox;

        private ProgressBar mProgressBar;
        private Grid mMainGrid;

        private TextEditor mTextView;
        //private TextEditor mDotTemplate_TextBox;

        private ListView mExcelListView;

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

            mExcelPath_btn = mMainGrid.FindName("excelpath_btn") as Button;
            mJsonPath_btn = mMainGrid.FindName("jsonpath_btn") as Button;
            mDotTemplateFilePath_btn = mMainGrid.FindName("templatefilepath_btn") as Button;
            mBegin_btn = mMainGrid.FindName("begin_btn") as Button;

            mExcelPath_TextBox = mMainGrid.FindName("exclepath_textbox") as TextBox;
            mJsonPath_TextBox = mMainGrid.FindName("jsonpath_textbox") as TextBox;
            mDotTemplateFilePath_TextBox = mMainGrid.FindName("templatefilepath_textbox") as TextBox;
            mSignsheet_Textbox = mMainGrid.FindName("signsheet_textbox") as TextBox;
            mEncryptionKey_Textbox = mMainGrid.FindName("encryptionkey_textbox") as TextBox;
            mEncryptionIV_Textbox = mMainGrid.FindName("encryptioniv_textbox") as TextBox;

            mMutilsheet_Checkbox = mMainGrid.FindName("mutilsheet_togglebtn") as ToggleButton;

            mTextView = mMainGrid.FindName("textview") as TextEditor;
            //mDotTemplate_TextBox = mMainGrid.FindName("dotcsfiletabitem") as TextEditor;

            mExcelListView = mMainGrid.FindName("excelfile_listview") as ListView;
            mExcelListView.ItemsSource = ListViweItemData;

            mDotcs_RadioBtn = mMainGrid.FindName("dotcs_radiobtn") as RadioButton;
            mDotts_RadioBtn = mMainGrid.FindName("dotts_radiobtn") as RadioButton;

            mJsonView_RadioBtn = mMainGrid.FindName("jsonradiobtn") as RadioButton;
            mTemplateView_RadioBtn = mMainGrid.FindName("templateradiobtn") as RadioButton;

            mExcelPath_TextBox.DataContext = ExcelPath;
            mJsonPath_TextBox.DataContext = JsonPath;
            mDotTemplateFilePath_TextBox.DataContext = TemplatePath;
            mSignsheet_Textbox.DataContext = SheetSign;
            mEncryptionKey_Textbox.DataContext = EncryptionKey;
            mEncryptionIV_Textbox.DataContext = EncryptionIV;

            if (mBegin_btn != null)
            {
                mBegin_btn.Content = "Begin";
                mBegin_btn.Click += (s, ee) => Button_ClickAsync(s, ee);
            }

            ComboBox cbox = mMainGrid.FindName("filternum_combobox") as ComboBox;
            cbox.SelectedIndex = HeadNum - 1;

            mProgressBar = mMainGrid.FindName("progressbar") as ProgressBar;
            mProgressBar.Value = 0;

            mMutilsheet_Checkbox.IsChecked = MultiSheet;
            mSignsheet_Textbox.Visibility = MultiSheet ? Visibility.Visible : Visibility.Hidden;

            EncryptionKey.Text = "";
            EncryptionIV.Text = "";

            mDotcs_RadioBtn.IsChecked = CSRadioBtnChecked;
            mDotts_RadioBtn.IsChecked = TSRadioBtnChecked;

            if (!CSRadioBtnChecked && !TSRadioBtnChecked) mDotcs_RadioBtn.IsChecked = true;

            mTextView.SyntaxHighlighting = JsonHighlighting;

            SetColor(Properties.Settings.Default.Color, Properties.Settings.Default.Theme);

            SetEncryptionUI(false);
            (mMainGrid.FindName("mutilsheet_label") as Label).IsEnabled = MultiSheet;
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
                mProgressBar.Maximum = mDataManages.FilesCount(HeadNum, isExistTemplatePath);
                mProgressBar.Value = 0;
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
                mJsonPath_TextBox.Text = mFolderDialog.FileName; // FolderDialog.SelectedPath.Trim();

            if (((Button)sender).Name.Equals("templatefilepath_btn"))
                mDotTemplateFilePath_TextBox.Text = mFolderDialog.FileName; // FolderDialog.SelectedPath.Trim();
        }

        private void ToggleButton_Checked(object sender, RoutedEventArgs e)
        {
            ToggleButton cb = sender as ToggleButton;
            switch (cb.Name)
            {
                case "encryption_togglebtn":
                    bool isChecked = (bool)cb.IsChecked;
                    mDataManages.CanEncryption = isChecked;
                    SetEncryptionUI(isChecked);
                    break;
                case "mutilsheet_togglebtn":
                    MultiSheet = (bool)cb.IsChecked;
                    Properties.Settings.Default.MultiSheet = MultiSheet;
                    mSignsheet_Textbox.Visibility = MultiSheet ? Visibility.Visible : Visibility.Hidden;
                    Label mutilsheetLabel = mMainGrid.FindName("mutilsheet_label") as Label;
                    mutilsheetLabel.IsEnabled = MultiSheet;
                    ShowFileList();
                    break;
            }
        }

        private void SetEncryptionUI(bool _isChecked)
        {
            ToggleButton encryptionBtn = mMainGrid.FindName("encryption_togglebtn") as ToggleButton;
            Label encryptionLabel = mMainGrid.FindName("encryption_label") as Label;

            ComboBox encryptionMode_ComboBox = mMainGrid.FindName("encryption_mode_combobox") as ComboBox;
            ComboBox encryptionPadding_ComboBox = mMainGrid.FindName("encryption_padding_combobox") as ComboBox;
            Label encryptionMode_Label = mMainGrid.FindName("encryption_mode_label") as Label;
            Label encryptionPadding_Label = mMainGrid.FindName("encryption_padding_label") as Label;

            List<ComboxEncryptionMode> list = new List<ComboxEncryptionMode>();
            int modeID = 0;
            foreach (var item in Enum.GetValues(typeof(EncryptionMode)))
            {
                modeID++;
                list.Add(new ComboxEncryptionMode(item.ToString(), modeID.ToString()));
            }
            encryptionMode_ComboBox.ItemsSource = list;

            List<ComboxEncryptionPadding> listPadding = new List<ComboxEncryptionPadding>();
            int paddingID = 0;
            foreach (var item in Enum.GetValues(typeof(EncryptionPadding)))
            {
                paddingID++;
                listPadding.Add(new ComboxEncryptionPadding(item.ToString(), paddingID.ToString()));
            }
            encryptionPadding_ComboBox.ItemsSource = listPadding;

            encryptionBtn.IsChecked = _isChecked;
            encryptionLabel.IsEnabled = _isChecked;
            encryptionMode_Label.IsEnabled = _isChecked;
            encryptionPadding_Label.IsEnabled = _isChecked;
            encryptionMode_ComboBox.IsEnabled = _isChecked;
            encryptionPadding_ComboBox.IsEnabled = _isChecked;
            mEncryptionKey_Textbox.IsEnabled = _isChecked;
            mEncryptionIV_Textbox.IsEnabled = _isChecked;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = sender as ComboBox;

            switch (cb.Name)
            {
                case "filternum_combobox":
                    HeadNum = cb.SelectedIndex + 1;
                    if (Properties.Settings.Default.HeadNum != HeadNum)
                        ShowFileList();

                    if (HeadNum > 1)
                    {
                        mDotcs_RadioBtn.Visibility = Visibility.Visible;
                        mDotts_RadioBtn.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        mDotcs_RadioBtn.Visibility = Visibility.Hidden;
                        mDotts_RadioBtn.Visibility = Visibility.Hidden;
                    }

                    Properties.Settings.Default.HeadNum = HeadNum;
                    break;
                case "encryption_mode_combobox":
                    mDataManages.Mode = (CipherMode)Enum.Parse(typeof(CipherMode), cb.SelectedIndex + 1 + "");
                    break;
                case "encryption_padding_combobox":
                    mDataManages.Padding = (PaddingMode)Enum.Parse(typeof(PaddingMode), cb.SelectedIndex + 1 + "");
                    break;
            }
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            switch (tb.Name)
            {
                case "exclepath_textbox":
                    ExcelPath.Text = tb.Text;
                    if (Directory.Exists(ExcelPath.Text) || string.IsNullOrEmpty(ExcelPath.Text))
                    {
                        Properties.Settings.Default.ExcelPath = ExcelPath.Text;
                        ShowFileList();
                    }
                    break;
                case "jsonpath_textbox":
                    JsonPath.Text = tb.Text;
                    if (Directory.Exists(JsonPath.Text) || string.IsNullOrEmpty(JsonPath.Text)) Properties.Settings.Default.JsonPath = JsonPath.Text;
                    break;
                case "templatefilepath_textbox":
                    TemplatePath.Text = tb.Text;
                    if (Directory.Exists(TemplatePath.Text) || string.IsNullOrEmpty(TemplatePath.Text)) Properties.Settings.Default.TemplatePath = TemplatePath.Text;
                    break;
                case "signsheet_textbox":
                    SheetSign.Text = tb.Text;
                    ShowFileList();
                    break;
                case "encryptionkey_textbox":
                    EncryptionKey.Text = tb.Text;
                    mDataManages.Key = tb.Text;
                    break;
                case "encryptioniv_textbox":
                    EncryptionIV.Text = tb.Text;
                    mDataManages.IV = tb.Text;
                    break;
            }
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
            switch (tb.Name)
            {
                case "exclepath_textbox":
                    ExcelPath.Text = fileName;
                    if (Directory.Exists(ExcelPath.Text))
                    {
                        Properties.Settings.Default.ExcelPath = ExcelPath.Text;
                        ShowFileList();
                    }
                    break;
                case "jsonpath_textbox":
                    JsonPath.Text = fileName;
                    if (Directory.Exists(JsonPath.Text))
                    {
                        Properties.Settings.Default.JsonPath = JsonPath.Text;
                    }
                    break;
                case "templatefilepath_textbox":
                    TemplatePath.Text = fileName;
                    if (Directory.Exists(TemplatePath.Text))
                    {
                        Properties.Settings.Default.TemplatePath = TemplatePath.Text;
                    }
                    break;
            }
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
            if (!(mExcelListView.SelectedItem is ListViewItemData obj)) return;
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
            if (JsonRadioBtnChecked) mTextView.Text = JsonData;
            if (TemplateRadioBtnChecked) mTextView.Text = TemplateData;
            //mDotTemplate_TextBox.Text = template;
        }

        private void Radiobtn_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton rb = sender as RadioButton;
            switch (rb.Name)
            {
                case "dotts_radiobtn":
                    Type = TemplateType.TS;
                    CSRadioBtnChecked = false;
                    TSRadioBtnChecked = true;
                    break;
                case "dotcs_radiobtn":
                    Type = TemplateType.CS;
                    CSRadioBtnChecked = true;
                    TSRadioBtnChecked = false;
                    break;
                case "jsonradiobtn":
                    TemplateRadioBtnChecked = false;
                    JsonRadioBtnChecked = true;
                    if (mTextView != null)
                    {
                        mTextView.SyntaxHighlighting = JsonHighlighting;
                        mTextView.Text = JsonData;
                    }
                    return;
                case "templateradiobtn":
                    TemplateRadioBtnChecked = true;
                    JsonRadioBtnChecked = false;
                    if (mTextView != null)
                    {
                        if (CSRadioBtnChecked)
                            mTextView.SyntaxHighlighting = CSHighlighting;
                        else if (TSRadioBtnChecked)
                            mTextView.SyntaxHighlighting = TSHighlighting;
                        mTextView.Text = TemplateData;
                    }
                    return;
            }
            Properties.Settings.Default.CSRadioBtnChecked = CSRadioBtnChecked;
            Properties.Settings.Default.TSRadioBtnChecked = TSRadioBtnChecked;
            ShowFileList();
        }

        private void ListBoxItem_MouseClick(object sender, RoutedEventArgs e)
        {
            string name = (sender as ListBoxItem).Name;
            if (name.Equals("Light") || name.Equals("Dark"))
                Properties.Settings.Default.Theme = name;
            else
                Properties.Settings.Default.Color = name;
            SetColor(Properties.Settings.Default.Color, Properties.Settings.Default.Theme);
        }

        // =======================================

        private void BgworkChange(object sender, ProgressChangedEventArgs e)
        {
            mProgressBar.Value += e.ProgressPercentage;
            if (mBegin_btn != null)
                mBegin_btn.Content = $"{FileName} Finished";

            if (mProgressBar.Value >= mProgressBar.Maximum)
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

            if (mBegin_btn != null)
            {
                mBegin_btn.Content = _isEnable ? "Begin" : "";
                mBegin_btn.IsEnabled = _isEnable;
            }
            if (mExcelPath_btn != null)
                mExcelPath_btn.IsEnabled = _isEnable;

            if (mJsonPath_btn != null)
                mJsonPath_btn.IsEnabled = _isEnable;

            if (mDotTemplateFilePath_btn != null)
                mDotTemplateFilePath_btn.IsEnabled = _isEnable;

            if (mExcelPath_TextBox != null)
                mExcelPath_TextBox.IsEnabled = _isEnable;

            if (mJsonPath_TextBox != null)
                mJsonPath_TextBox.IsEnabled = _isEnable;

            if (mDotTemplateFilePath_TextBox != null)
                mDotTemplateFilePath_TextBox.IsEnabled = _isEnable;
        }

        private void DoShowFileList(object sender, DoWorkEventArgs e)
        {
            if (Directory.Exists(ExcelPath.Text) && mExcelListView != null)
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
            mTextView?.Clear();
            ListViweItemData?.Clear();
            mDataManages?.ClearData();
            JsonRadioBtnChecked = true;
            TemplateRadioBtnChecked = false;
            mJsonView_RadioBtn.IsChecked = JsonRadioBtnChecked;
            mTemplateView_RadioBtn.IsChecked = TemplateRadioBtnChecked;
            //this.Width = this.MinWidth;

            if (Directory.Exists(ExcelPath.Text) && mExcelListView != null)
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

            if (mProgressBar != null)
                mProgressBar.Value = 0;
        }

        private void SetColor(string _colorName, string _theme)
        {
            Color primaryColor = SwatchHelper.Lookup[MaterialDesignColor.DeepOrange];
            Color accentColor = SwatchHelper.Lookup[MaterialDesignColor.DeepOrange];
            IBaseTheme theme = new MaterialDesignDarkTheme();
            switch (_colorName)
            {
                case "Light":
                    theme = new MaterialDesignLightTheme();
                    break;
                case "Dark":
                    theme = new MaterialDesignDarkTheme();
                    break;
                case "Yellow":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Yellow];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Yellow];
                    break;
                case "Amber":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Amber];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Amber];
                    break;
                case "DeepOrange":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.DeepOrange];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.DeepOrange];
                    break;
                case "Lightblue":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.LightBlue];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.LightBlue];
                    break;
                case "Teal":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Teal];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Teal];
                    break;
                case "Cyan":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Cyan];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Cyan];
                    break;
                case "Pink":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Pink];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Pink];
                    break;
                case "Green":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Green];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Green];
                    break;
                case "DeepPurple":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.DeepPurple];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.DeepPurple];
                    break;
                case "Indigo":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Indigo];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Indigo];
                    break;
                case "LightGreen":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.LightGreen];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.LightGreen];
                    break;
                case "Blue":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Blue];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Blue];
                    break;
                case "Lime":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Lime];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Lime];
                    break;
                case "Red":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Red];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Red];
                    break;
                case "Orange":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Orange];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Orange];
                    break;
                case "Purple":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Purple];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Purple];
                    break;
                case "BlueGrey":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.BlueGrey];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.BlueGrey];
                    break;
                case "Grey":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Grey];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Grey];
                    break;
                case "Brown":
                    primaryColor = SwatchHelper.Lookup[MaterialDesignColor.Brown];
                    accentColor = SwatchHelper.Lookup[MaterialDesignColor.Brown];
                    break;
            }

            switch (_theme)
            {
                case "Light":
                    theme = new MaterialDesignLightTheme();
                    mTextView.Foreground = Brushes.Black;
                    break;
                case "Dark":
                    theme = new MaterialDesignDarkTheme();
                    mTextView.Foreground = Brushes.White;
                    break;
            }

            ITheme themes = Theme.Create(theme, primaryColor, accentColor);
            Resources.SetTheme(themes);
        }
    }
}