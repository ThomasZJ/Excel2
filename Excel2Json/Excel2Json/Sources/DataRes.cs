using System;
using System.IO;

namespace Excel2
{
    public enum TemplateType
    {
        MIN = 0,
        CS,
        TS,
        MAX
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
        public int ID { get; set; }
        public string Name { get; set; }

        public ComboxEncryptionMode(int _id, string _name)
        {
            ID = _id;
            Name = _name;
        }
    }
}
