using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using ExcelDataReader;
using Newtonsoft.Json;

namespace Excel2
{
    /// <summary>
    /// 导出文件类型
    /// </summary>

    internal class DataManages
    {
        /// <summary>
        /// Excel Data
        /// </summary>
        public Dictionary<string, DataSet> ExcelData { get; }

        /// <summary>
        /// Json Data
        /// </summary>
        public Dictionary<string, string> JsonData { get; }

        /// <summary>
        /// Template Data
        /// </summary>
        public Dictionary<string, string> TemplateData { get; }

        public CipherMode Mode { set; get; }
        public PaddingMode Padding { set; get; }

        public string Key { get; set; }
        public string IV { get; set; }
        public bool CanEncryption { get; set; }

        public DataManages()
        {
            ExcelData = new Dictionary<string, DataSet>();
            JsonData = new Dictionary<string, string>();
            TemplateData = new Dictionary<string, string>();
        }

        public DataSet ReadExcel(FileInfo _file)
        {
            DataSet dataSet = null;
            using (var stream = File.Open(_file.FullName, FileMode.Open, FileAccess.Read))
            {
                using (var excelReader = ExcelReaderFactory.CreateReader(stream))
                {
                    string name = _file.Name.Split('.')[0];
                    dataSet = excelReader.AsDataSet();
                    if (!ExcelData.ContainsKey(name))
                        ExcelData.Add(name, dataSet);
                    excelReader.Close();
                }
                stream.Close();
            }
            return dataSet;
        }

        public void ExportJson(FileInfo _file, int _headNum, bool _isMutiple, string _sheetSign = "")
        {
            string name = _file.Name.Split('.')[0];
            if (!ExcelData.ContainsKey(name)) return;
            DataSet excelData = ExcelData[name];
            if (excelData == null) return;

            var jsonSettings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented
            };
            if (!_isMutiple)
            {
                DataTable dataTabale = excelData.Tables[0];
                if (dataTabale.Rows.Count > 0 && dataTabale.Columns.Count > 0)
                {
                    object sheetValue = ConvertSheetToArray(dataTabale, _headNum);
                    //object sheetValue = ConvertSheetToDict(dataTabale, _headNum);
                    string context = JsonConvert.SerializeObject(sheetValue, jsonSettings);
                    if (!JsonData.ContainsKey(name))
                        JsonData.Add(name, context);
                }
            }
            else
            {
                foreach (DataTable item in excelData.Tables)
                {
                    if (!string.IsNullOrEmpty(_sheetSign) && item.TableName.Contains(_sheetSign)) continue;
                    if (item.Rows.Count > 0 && item.Columns.Count > 0)
                    {
                        object sheetValue = ConvertSheetToArray(item, _headNum);
                        string jsonContext = JsonConvert.SerializeObject(sheetValue, jsonSettings);
                        if (!JsonData.ContainsKey(/*name + */item.TableName))
                            JsonData.Add(/*name + */item.TableName, jsonContext);
                    }
                }
            }
        }

        public void ExportTemplate(FileInfo _file, int _headNum, bool _isMutiple, TemplateType _template, string _sheetSign = "")
        {
            string name = _file.Name.Split('.')[0];
            if (!ExcelData.ContainsKey(name)) return;
            DataSet excelData = ExcelData[name];
            if (excelData == null) return;

            if (!_isMutiple)
            {
                DataTable dataTabale = excelData.Tables[0];
                if (dataTabale.Rows.Count > 0 && dataTabale.Columns.Count > 0)
                {
                    string tmp = "";
                    switch (_template)
                    {
                        case TemplateType.CS:
                            tmp = (new CSDefineGenerator().CSGenerator(name, _headNum, dataTabale));
                            break;
                        case TemplateType.TS:
                            tmp = (new TypeScriptGenerator().TSGenerator(name, _headNum, dataTabale));
                            break;
                    }
                    if (!TemplateData.ContainsKey(name))
                        TemplateData.Add(name, tmp);
                }
            }
            else
            {
                foreach (DataTable item in excelData.Tables)
                {
                    if (!string.IsNullOrEmpty(_sheetSign) && item.TableName.Contains(_sheetSign)) continue;
                    if (item.Rows.Count > 0 && item.Columns.Count > 0)
                    {
                        string tmp = "";
                        switch (_template)
                        {
                            case TemplateType.CS:
                                tmp = (new CSDefineGenerator().CSGenerator(/*name + */item.TableName, _headNum, item));
                                break;
                            case TemplateType.TS:
                                tmp = (new TypeScriptGenerator().TSGenerator(/*name + */item.TableName, _headNum, item));
                                break;
                        }
                        if (!TemplateData.ContainsKey(/*name + */item.TableName))
                            TemplateData.Add(/*name + */item.TableName, tmp);
                    }
                }

            }
        }

        /// <summary>
        /// 获取文件数量
        /// </summary>
        /// <param name="_headNum">excel 文件头</param>
        /// <returns></returns>
        public int FilesCount(int _headNum, bool _isExistTemplatePath)
        {
            if (_headNum > 1 && _isExistTemplatePath)
                return JsonData.Count + TemplateData.Count;
            else
                return JsonData.Count;

        }

        /// <summary>
        /// 清理数据文件
        /// </summary>
        internal void ClearData()
        {
            ExcelData.Clear();
            JsonData.Clear();
            TemplateData.Clear();
        }

        /// <summary>
        /// Save the json data and template data to files
        /// </summary>
        /// <param name="_jsonPath">the json files path</param>
        /// <param name="templatePath">the template files path</param>
        /// <param name="_headNum">the excel head number</param>
        /// <param name="callback"></param>
        public void SaveFiles(string _jsonPath, string templatePath, int _headNum, TemplateType _type, Action<double, string> callback)
        {
            if (Directory.Exists(_jsonPath))
            {
                int time = 500 / JsonData.Count;
                foreach (var item in JsonData)
                {
                    string fileName = _jsonPath + "\\" + item.Key + ".json";
                    string jsonData = item.Value;

                    if (CanEncryption)
                        jsonData = DesEncrypt(Key, IV, jsonData, Mode, Padding);

                    using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        using (TextWriter writer = new StreamWriter(file, new UTF8Encoding(false)))
                        {
                            writer.Write(jsonData);
                        }
                        file.Close();
                        callback(1, item.Key + ".json");
                    }
                    Thread.Sleep(time);
                }
            }

            if (Directory.Exists(templatePath) && _headNum > 1 && _type != TemplateType.MIN)
            {
                int time = 500 / TemplateData.Count;
                foreach (var item in TemplateData)
                {
                    string suffix = "";
                    switch (_type)
                    {
                        case TemplateType.CS:
                            suffix = ".cs";
                            break;
                        case TemplateType.TS:
                            suffix = ".ts";
                            break;
                        default:
                            break;
                    }
                    string fileName = templatePath + "\\" + item.Key + suffix;
                    string templateData = item.Value;

                    //if (CanEncryption)
                    //    jsonData = DesEncrypt(Key, IV, jsonData, Mode, Padding);

                    using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        using (TextWriter writer = new StreamWriter(file, new UTF8Encoding(false)))
                        {
                            writer.Write(templateData);
                        }
                        file.Close();
                        callback(1, item.Key + suffix);
                    }
                    Thread.Sleep(time);
                }
            }
        }

        public void SaveFile(string _jsonPath, string templatePath, int _headNum, TemplateType _type, string _fileName, Action<string> callback)
        {
            if (Directory.Exists(_jsonPath))
            {
                if (JsonData.ContainsKey(_fileName))
                {
                    string fileName = _jsonPath + "\\" + _fileName + ".json"; ;
                    string jsonData = JsonData[_fileName];
                    if (CanEncryption)
                        jsonData = DesEncrypt(Key, IV, jsonData, Mode, Padding);
                    using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        using (TextWriter writer = new StreamWriter(file, new UTF8Encoding(false)))
                        {
                            writer.Write(jsonData);
                        }
                        file.Close();
                        callback?.Invoke(_fileName + ".json");
                    }
                }
            }

            if (Directory.Exists(templatePath) && _headNum > 1 && _type != TemplateType.MIN)
            {
                if (TemplateData.ContainsKey(_fileName))
                {
                    string suffix = "";
                    switch (_type)
                    {
                        case TemplateType.CS:
                            suffix = ".cs";
                            break;
                        case TemplateType.TS:
                            suffix = ".ts";
                            break;
                        default:
                            break;
                    }
                    string fileName = templatePath + "\\" + _fileName + suffix;
                    string templateData = TemplateData[_fileName];

                    //if (CanEncryption)
                    //    templateData = DesEncrypt(Key, IV, templateData, Mode, Padding);

                    using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        using (TextWriter writer = new StreamWriter(file, new UTF8Encoding(false)))
                        {
                            writer.Write(templateData);
                        }
                        file.Close();
                        callback?.Invoke(_fileName + suffix);
                    }
                }
            }
        }

        /// <summary>
        /// Change the excel data to array
        /// </summary>
        /// <param name="_dt">DataTable </param>
        /// <param name="_firstDataRow">the excel's first line except head (不包含表头的第一行) </param>
        /// <returns></returns>
        private object ConvertSheetToArray(DataTable _dt, int _firstDataRow)
        {
            List<object> values = new List<object>();
            for (int i = _firstDataRow; i < _dt.Rows.Count; i++)
            {
                DataRow row = _dt.Rows[i];
                values.Add(ConvertRowToDict(_dt, row, _firstDataRow));
            }
            return values;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object ConvertSheetToDict(DataTable _dt, int _firstDataRow)
        {
            Dictionary<string, object> importData =
                new Dictionary<string, object>();

            int firstDataRow = 0;
            for (int i = firstDataRow; i < _dt.Rows.Count; i++)
            {
                DataRow row = _dt.Rows[i];
                string ID = row[_dt.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = ConvertRowToDict(_dt, row, _firstDataRow);
                rowObject[ID] = ID;
                importData[ID] = rowObject;
            }

            return importData;
        }

        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> ConvertRowToDict(DataTable _dt, DataRow row, int firstDataRow)
        {
            var rowData = new Dictionary<string, object>();
            foreach (DataColumn column in _dt.Columns)
            {
                object value = row[column];

                if (value.GetType() == typeof(System.DBNull))
                {
                    value = GetColumnDefault(_dt, column, firstDataRow);
                }
                else if (firstDataRow > 1)
                {
                    try
                    {
                        switch (_dt.Rows[1][column])
                        {
                            case "int":
                                if (value.GetType() == typeof(double))
                                { // 去掉数值字段的“.0”
                                    double num = (double)value;
                                    if ((int)num == num)
                                        value = (int)num;
                                }
                                else
                                    value = int.Parse(value.ToString());
                                break;
                            case "float":
                                value = float.Parse(value.ToString());
                                break;
                            case "double":
                                value = double.Parse(value.ToString());
                                break;
                            default:
                                value = value.ToString();
                                break;
                        }
                    }
                    catch (FormatException)
                    {
                        // throw new FormatException("has not correct format");
                        value = value.ToString();
                    }
                }
                // 表头自动转换成小写
                //if (lowcase)
                //    fieldName = fieldName.ToLower();
                string fieldName = _dt.Rows[0][column].ToString();
                rowData[fieldName] = value;
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object GetColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
        {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull))
                {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }

        public string DesEncrypt(string _key, string _iv, string _orgText, CipherMode _mode, PaddingMode padding)
        {
            if (_key.Length < 8)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(_key);
                for (int i = _key.Length; i < 8; i++)
                {
                    sb.Append("0");
                }
                _key = sb.ToString();
            }
            byte[] inputByteArray = Encoding.UTF8.GetBytes(_orgText);
            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            des.Mode = _mode;
            des.Padding = padding;
            des.Key = ASCIIEncoding.ASCII.GetBytes(_key);
            des.IV = ASCIIEncoding.ASCII.GetBytes(_key);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();
            StringBuilder ret = new StringBuilder();
            foreach (byte b in ms.ToArray())
            {
                ret.AppendFormat("{0:X2}", b);
            }
            string encryptStr = ret.ToString();  // Encoding.UTF8.GetString(ms.ToArray());
            return encryptStr;
        }
    }
}
