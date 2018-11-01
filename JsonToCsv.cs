using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LitJson;
using System.Text.RegularExpressions;
using System.Xml;

namespace excel2json
{
    class JsonToCsv
    {
        public static void Process(string srcFilename, string targetFilename, string settingFilename)
        {
            List<KeyValuePair<string, string[]>> columnList = readColumnSetting(settingFilename);
            
            List<List<string>> data = new List<List<string>>();
            List<string[]> columnPathList = new List<string[]>();

            List<string> title = new List<string>();
            foreach(KeyValuePair<string, string[]> columnPair in columnList)
            {
                columnPathList.Add(columnPair.Value);
                title.Add(columnPair.Key);
            }
            data.Add(title);

            JsonData pejd = JsonMapper.ToObject(System.IO.File.ReadAllText(srcFilename));
            
            for (int j = 0; j < pejd.Count; j++)
            {
                List<string> row = new List<string>();
                for (int i = 0; i < columnPathList.Count; i++)
                {
                    ProcessColumn(row, columnPathList[i], pejd[j]);
                }
                data.Add(row);
            }

            System.IO.FileStream fs = new System.IO.FileStream(targetFilename, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, Encoding.Default);
            foreach(List<string> i in data)
            {
                StringBuilder sb = new StringBuilder();

                foreach(string s in i)
                {
                    sb.Append(s);
                    sb.Append(",");
                }
                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
                string l = sb.ToString();
                l = Encoding.Default.GetString(Encoding.Convert(Encoding.UTF8, Encoding.Default, Encoding.UTF8.GetBytes(l)));
                sw.Write(l);
                
            }

            sw.Close();
            fs.Close();
        }
        public static void ProcessColumn(List<string> data, string[] columnPath, JsonData playerExtent)
        {
            bool hasValue = true;
            JsonData value = playerExtent;
            for (int i = 0; i < columnPath.Length; i++)
            {
                string isArray = Regex.Match(columnPath[i], @"\[\d+\]").Value;
                if (isArray.Length != 0)
                {
                    string arrayStr = Regex.Match(isArray, @"\d+").Value;
                    int arrayId = int.Parse(arrayStr);
                    if (value.IsArray)
                        value = value[arrayId];
                    else
                    {
                        hasValue = false;
                        break;
                    }
                }
                else
                {
                    if (!value.IsObject || !value.Keys.Contains(columnPath[i]))
                    {
                        hasValue = false;
                        break;
                    }
                    value = value[columnPath[i]];
                }
            }
            if (hasValue)
            {
                if (value.IsString)
                    data.Add(value.GetString());
                else if (value.IsInt)
                    data.Add(value.GetInt().ToString());
                else if (value.IsDouble)
                    data.Add(value.GetDouble().ToString());
                else
                    data.Add("$error$");
            }
            else
            {
                data.Add("");
            }
        }
        static List<KeyValuePair<string, string[]>> readColumnSetting(string filename)
        {
            List<KeyValuePair<string, string[]>> rt = new List<KeyValuePair<string, string[]>>();
            XmlDocument doc = new XmlDocument();
            doc.Load(filename);

            XmlNodeList columnList = doc.DocumentElement.SelectNodes("column");
            for (int i = 0; i < columnList.Count; i++)
            {
                string colName = columnList.Item(i).Attributes["name"].Value;
                string colPath = columnList.Item(i).Attributes["path"].Value;

                string[] columnPaths = colPath.Split(new char[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);

                KeyValuePair<string, string[]> col = new KeyValuePair<string, string[]>(colName, columnPaths);
                rt.Add(col);
            }
            return rt;
        }
    }
}
