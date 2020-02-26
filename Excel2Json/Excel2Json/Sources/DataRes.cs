using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;

namespace Excel2
{
    public enum TemplateType
    {
        MIN, CS, TS,
    }

    public enum EncryptionMode
    {
        CBC, ECB, OFB, CFB
    }

    public enum EncryptionPadding
    {
        None, PKCS7, Zeros, ANSIX923, ISO10126
    }

    class DataRes
    {
    }

    public class ListViewItemData
    {
        public string ID { get; set; }
        public string FileName { get; set; }
        public string FullName { get; set; }

        public FileInfo FileInfo { get; private set; }

        public ListViewItemData(string _id, FileInfo _file)
        {
            ID = _id;
            FileName = _file.Name;
            FullName = _file.FullName;
            FileInfo = _file;
        }
    }

    public class ComboxEncryptionMode
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public ComboxEncryptionMode(string _name, string _value)
        {
            Name = _name;
            Value = _value;
        }
    }

    public class ComboxEncryptionPadding
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public ComboxEncryptionPadding(string _name, string _value)
        {
            Name = _name;
            Value = _value;
        }
    }

    public class TextBoxData : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string mText;
        public string Text
        {
            get
            {
                return mText;
            }
            set
            {
                mText = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Text"));
            }
        }

    }
}
