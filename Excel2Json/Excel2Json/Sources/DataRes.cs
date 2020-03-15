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
    public enum Themes
    {
        Light,
        Dark,
        Yellow,
        Amber,
        DeepOrange,
        LightBlue,
        Teal,
        Cyan,
        Pink,
        Green,
        DeepPurple,
        Indigo,
        LightGreen,
        Blue,
        Lime,
        Red,
        Orange,
        Purple,
        BlueGrey,
        Grey,
        Brown
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

    public class ThemesListBoxItem
    {
        public string Name { get; set; }
        public ThemesListBoxItem(string _name)
        {
            Name = _name;
        }
    }
}
