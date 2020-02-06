using System.Collections.Generic;
/*
 * ******************************
 * File: TemplateGenerator.cs
 * File Created: Tuesday, 14th January 2020 16:36:36
 * Author: Thomas (Thomas_Joker@outlook.com)
 * ------------------------------
 * Last Modified: Tuesday, 14th January 2020 16:36:36
 * Modified By: Thomas (Thomas_Joker@outlook.com>)
 * ------------------------------
 * Copyright(c) 2019 - 2020 *.net, *.net
 * ******************************
 */
using System.Text;

namespace Excel2Json
{
    internal class TemplateGenerator
    {
        public enum TemplateType
        {
            CSharp,
            CPP,
            Go,
            Java,
            Kotlin
        }

        public struct TemplateField
        {
            public string name;
            public string type;
            public string comment;
        }

        // public TemplateGenerator(TemplateType _type, TemplateField _field, string _name)
        // {
        //     switch (_type)
        //     {
        //         case TemplateType.CSharp:
        //             GeneratorCSharp(_field, _name);
        //             break;
        //         default: break;
        //     }
        // }

        public Dictionary<string, string> GeneratorCSharp(TemplateField _field, string _name)
        {
            Dictionary<string, string> code = new Dictionary<string, string>();
            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("//");
            sb.AppendLine("// Auto Generated Code By Excel2json");
            sb.AppendLine("//");
            sb.AppendLine();
            sb.AppendFormat("// Generate From {0}.xlsx", _name);
            sb.AppendLine();
            sb.AppendFormat("public class {0}\r\n{{", _name);
            sb.AppendLine();

            sb.AppendFormat("\t/// <summary>\n");
            sb.AppendFormat("\t/// {0}\n", _field.comment);
            sb.AppendFormat("\t/// </summary>\n");
            sb.AppendFormat("\tpublic {0} {1}; ", _field.type, _field.name);
            sb.AppendLine();

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            code.Add(_name, sb.ToString());
            return code;
        }
    }
}
