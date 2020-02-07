using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using Microsoft.WindowsAPICodePack.Dialogs;

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

        private RadioButton mDotcs_RadioBtn;
        private RadioButton mDotts_RadioBtn;

        private CheckBox mMutilsheet_Checkbox;

        private ProgressBar mProgressBar;
        private Grid mMainGrid;

        private TextEditor mTextView;
        //private TextEditor mDotTemplate_TextBox;

        private ListView mExcelListView;

        private readonly DataManages mDataManages;
        private ObservableCollection<ListViewItemData> ListViweItemData;

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

        /// <summary>
        /// excel文件路径 
        /// </summary>
        private string ExcelPath { get; set; } = "";

        /// <summary>
        /// 导出json路径
        /// </summary>
        private string JsonPath { get; set; } = "";

        /// <summary>
        /// Template Classes path
        /// </summary>
        private string TemplatePath { get; set; } = "";

        /// <summary>
        /// Filter key word for sheet
        /// </summary>
        public string SheetSign { private set; get; }

        private string JsonData { get; set; }

        private string TemplateData { get; set; }



        private BackgroundWorker mBgworker;
        private DoWorkEventHandler mDoWorkEventHandler;
        private ProgressChangedEventHandler mProgressChangedEventHandler;
        private CommonOpenFileDialog mFolderDialog;

        private BackgroundWorker mBgShowFileList;
        private DoWorkEventHandler mDoShowFileHandler;

        private IHighlightingDefinition JsonHighlighting;
        private IHighlightingDefinition CSHighlighting;
        private IHighlightingDefinition TSHighlighting;

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

            ExcelPath = Properties.Settings.Default.ExcelPath;
            JsonPath = Properties.Settings.Default.JsonPath;
            TemplatePath = Properties.Settings.Default.TemplatePath;
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

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
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

            mMutilsheet_Checkbox = mMainGrid.FindName("mutilsheet_checkbox") as CheckBox;

            mTextView = mMainGrid.FindName("textview") as TextEditor;
            //mDotTemplate_TextBox = mMainGrid.FindName("dotcsfiletabitem") as TextEditor;

            mExcelListView = mMainGrid.FindName("excelfile_listview") as ListView;
            mExcelListView.ItemsSource = ListViweItemData;

            mDotcs_RadioBtn = mMainGrid.FindName("dotcs_radiobtn") as RadioButton;
            mDotts_RadioBtn = mMainGrid.FindName("dotts_radiobtn") as RadioButton;

            //mDotJsonView_RadioBtn = mMainGrid.FindName("jsonradiobtn") as RadioButton;
            //mDotFilesView_RadioBtn = mMainGrid.FindName("dotfileradiobtn") as RadioButton;

            if (mBegin_btn != null)
            {
                mBegin_btn.Content = "Begin";
                mBegin_btn.Click += (s, ee) => Button_ClickAsync(s, ee);
            }

            if (!string.IsNullOrEmpty(ExcelPath))
                mExcelPath_TextBox.Text = ExcelPath;

            if (!string.IsNullOrEmpty(JsonPath))
                mJsonPath_TextBox.Text = JsonPath;

            if (!string.IsNullOrEmpty(TemplatePath))
                mDotTemplateFilePath_TextBox.Text = TemplatePath;

            ComboBox cbox = mMainGrid.FindName("filternum_combobox") as ComboBox;
            cbox.SelectedIndex = HeadNum - 1;

            mProgressBar = mMainGrid.FindName("progressbar") as ProgressBar;
            mProgressBar.Value = 0;

            mMutilsheet_Checkbox.IsChecked = MultiSheet;
            mSignsheet_Textbox.Visibility = MultiSheet ? Visibility.Visible : Visibility.Hidden;

            mDotcs_RadioBtn.IsChecked = CSRadioBtnChecked;
            mDotts_RadioBtn.IsChecked = TSRadioBtnChecked;

            if (!CSRadioBtnChecked && !TSRadioBtnChecked) mDotcs_RadioBtn.IsChecked = true;

            mTextView.SyntaxHighlighting = JsonHighlighting;
        }

        private void Button_ClickAsync(object sender, RoutedEventArgs e)
        {
            bool isExistExclePath = true;
            bool isExistJsonPath = true;
            bool isExistTemplatePath = true;


            if (!Directory.Exists(ExcelPath))
                isExistExclePath = false;
            if (!Directory.Exists(JsonPath))
                isExistJsonPath = false;
            if (!Directory.Exists(TemplatePath))
                isExistTemplatePath = false;

            if (!isExistExclePath || !isExistJsonPath)
            {
                if (!isExistExclePath)
                    MessageBox.Show("ExcelPath does not exist!", "Directory not exist!");

                else if (!isExistJsonPath)
                    MessageBox.Show("JsonPath does not exist!", "Directory not exist");

                SetUIStates(true);
            }
            else
            {
                //TODO
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
                mExcelPath_TextBox.Text = mFolderDialog.FileName;// FolderDialog.SelectedPath.Trim();

            if (((Button)sender).Name.Equals("jsonpath_btn"))
                mJsonPath_TextBox.Text = mFolderDialog.FileName; // FolderDialog.SelectedPath.Trim();

            if (((Button)sender).Name.Equals("templatefilepath_btn"))
                mDotTemplateFilePath_TextBox.Text = mFolderDialog.FileName; // FolderDialog.SelectedPath.Trim();
        }

        private void Multi_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox cb = sender as CheckBox;

            if (cb.Name.Equals("mutilsheet_checkbox"))
            {
                MultiSheet = (bool)cb.IsChecked;
                Properties.Settings.Default.MultiSheet = MultiSheet;
                mSignsheet_Textbox.Visibility = MultiSheet ? Visibility.Visible : Visibility.Hidden;

                ShowFileList();
            }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = sender as ComboBox;
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
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            switch (tb.Name)
            {
                case "exclepath_textbox":
                    ExcelPath = tb.Text;
                    if (Directory.Exists(ExcelPath) || string.IsNullOrEmpty(ExcelPath))
                    {
                        Properties.Settings.Default.ExcelPath = ExcelPath;
                        ShowFileList();
                    }
                    break;
                case "jsonpath_textbox":
                    JsonPath = tb.Text;
                    if (Directory.Exists(JsonPath) || string.IsNullOrEmpty(JsonPath)) Properties.Settings.Default.JsonPath = JsonPath;
                    break;
                case "templatefilepath_textbox":
                    TemplatePath = tb.Text;
                    if (Directory.Exists(TemplatePath) || string.IsNullOrEmpty(TemplatePath)) Properties.Settings.Default.TemplatePath = TemplatePath;
                    break;
                case "signsheet_textbox":
                    SheetSign = tb.Text;
                    ShowFileList();
                    break;
            }

            //Trace.WriteLine(ExcelPath);
            //Trace.WriteLine(JsonPath);
            //Trace.WriteLine(TemplatePath);
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
                    ExcelPath = fileName;
                    if (Directory.Exists(ExcelPath))
                    {
                        Properties.Settings.Default.ExcelPath = ExcelPath;
                        ShowFileList();
                    }
                    break;
                case "jsonpath_textbox":
                    JsonPath = fileName;
                    if (Directory.Exists(JsonPath))
                    {
                        Properties.Settings.Default.JsonPath = JsonPath;
                    }
                    break;
                case "templatefilepath_textbox":
                    TemplatePath = fileName;
                    if (Directory.Exists(TemplatePath))
                    {
                        Properties.Settings.Default.TemplatePath = TemplatePath;
                    }
                    break;
            }
        }

        private void Textbox_LostFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            TextBox cb = sender as TextBox;
            if (cb.Name.Equals("exclepath_textbox")) ShowFileList();
            else if (cb.Name.Equals("jsonpath_textbox")) { }
            else if (cb.Name.Equals("templatefilepath_textbox")) { }
            else if (cb.Name.Equals("signsheet_textbox")) ShowFileList();
        }

        private void ListView_MouseClick(object sender, RoutedEventArgs e)
        {
            if (!(mExcelListView.SelectedItem is ListViewItemData obj)) return;
            //this.Width = 1075;
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
            if (TemplateRadioBtnChecked) mTextView.Text = TemplateData;
            if (JsonRadioBtnChecked) mTextView.Text = JsonData;
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
                case "dotfileradiobtn":
                    TemplateRadioBtnChecked = true;
                    JsonRadioBtnChecked = false;
                    if (mTextView != null)
                    {
                        if (CSRadioBtnChecked)
                            mTextView.SyntaxHighlighting = CSHighlighting;
                        else if (TSRadioBtnChecked)
                            mTextView.SyntaxHighlighting = TSHighlighting;

                        mTextView.Foreground = Brushes.White;
                        mTextView.Text = TemplateData;
                    }
                    return;
            }
            Properties.Settings.Default.CSRadioBtnChecked = CSRadioBtnChecked;
            Properties.Settings.Default.TSRadioBtnChecked = TSRadioBtnChecked;
            ShowFileList();
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
            if (Directory.Exists(JsonPath))
            {
                mDataManages.SaveFiles(JsonPath, TemplatePath, HeadNum, Type, (d, v) =>
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
            if (Directory.Exists(ExcelPath) && mExcelListView != null)
            {
                DirectoryInfo TheFolder = new DirectoryInfo(ExcelPath);
                foreach (var item in TheFolder.GetFiles())
                {
                    if (item.Extension.Equals(".xlsx") || item.Extension.Equals(".xls"))
                    {
                        mDataManages.ReadExcel(item);
                        mDataManages.ExportJson(item, HeadNum, MultiSheet, SheetSign);
                        mDataManages.ExportTemplate(item, HeadNum, MultiSheet, Type, SheetSign);
                    }
                }
            }
        }

        private void ShowFileList()
        {
            mTextView?.Clear();
            ListViweItemData?.Clear();
            mDataManages?.ClearData();
            //this.Width = this.MinWidth;

            if (Directory.Exists(ExcelPath) && mExcelListView != null)
            {
                DirectoryInfo TheFolder = new DirectoryInfo(ExcelPath);
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
    }
}

