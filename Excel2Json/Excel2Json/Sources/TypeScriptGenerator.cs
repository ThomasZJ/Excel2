using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Excel2
{
    internal class TypeScriptGenerator
    {
        private struct FieldDef
        {
            public string name;
            public string type;
            public string comment;
        }

        private string GeneratorDotTS(string _name, List<FieldDef> _fieldList)
        {
            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(@"//");
            sb.AppendLine(@"// Auto Generated Code By Excel2");
            sb.AppendLine(@"//");
            sb.AppendLine();
            sb.AppendFormat("// Generate From {0}.xlsx", _name);
            sb.AppendLine();
            sb.AppendFormat("interface {0}\r\n{{", _name);
            sb.AppendLine();

            foreach (FieldDef field in _fieldList)
            {
                sb.AppendFormat("\t/// <summary>\n");
                if (!string.IsNullOrEmpty(field.comment))
                    sb.AppendFormat("\t/// {0}\n", field.comment);
                sb.AppendFormat("\t/// </summary>\n");
                sb.AppendFormat("\tpublic {0}:{1}; ", field.name, field.type);
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine(@"// End of Auto Generated Code");

            return sb.ToString();
        }

        public string TSGenerator(string fileName, int _headNum, DataTable _excle)
        {
            if (_excle.Rows.Count < 1)
                return null;
            List<FieldDef> fieldList = new List<FieldDef>();

            if (_headNum > 1)
            {
                DataRow name = _excle.Rows[0];
                DataRow type = _excle.Rows[1];
                DataRow comment = null;
                if (_headNum >= 3)
                    comment = _excle.Rows[2];

                for (int k = 0; k < name.ItemArray.Length; k++)
                {
                    FieldDef field;
                    field.name = name[k].ToString();

                    string tmp = type[k].ToString();
                    switch (tmp)
                    {
                        case "int":
                        case "float":
                        case "double":
                            tmp = "number";
                            break;
                        case "string":
                            tmp = "string";
                            break;
                        case "boolean":
                            tmp = "boolean";
                            break;
                    }

                    if (string.IsNullOrEmpty(tmp)) tmp = "any";
                    field.type = tmp;
                    if (comment != null)
                        field.comment = comment[k].ToString();
                    else
                        field.comment = "";
                    fieldList.Add(field);
                }
                return GeneratorDotTS(fileName, fieldList);
            }
            return null;
        }
    }
}
