using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace excel2json
{
    class Program
    {
        static void Main1(string[] args)
        {
            Excel.Application app = new Excel.Application();

            string targetPath;
            string srcPath = "";
            string settingFile = Directory.GetCurrentDirectory() + "\\columnSetting.xml";
            bool allSheet = true;
            bool needDataType = true;
            if (args.Length == 1)
            {
                targetPath = args[0];
            }
            else if (args.Length == 2)
            {
                targetPath = args[0];
                settingFile = args[1];
            }
            else if (args.Length == 3)
            {
                targetPath = args[0];
                allSheet = args[1] == "allSheet" ? true : false;
                needDataType = args[2] == "needDataType" ? true : false;
            }
            else if (args.Length == 4)
            {
                srcPath = args[0];
                targetPath = args[1];
                allSheet = args[2] == "allSheet" ? true : false;
                needDataType = args[3] == "needDataType" ? true : false;
            }
            else
            {
                Console.WriteLine("1, excel2json excelFile");
                Console.WriteLine("2, excel2json outDir allSheet needDataType");
                Console.WriteLine("3, excel2json srcDir outDir allSheet needDataType");
                Console.WriteLine("4, excel2json jsonFile");

                return;
            }

            bool needPressKey = true;

            if (File.Exists(targetPath))
            {
                if (Path.GetExtension(targetPath) == ".json")
                {
                    Console.WriteLine("                     --------  {0} turn to csv --------", Path.GetFileNameWithoutExtension(targetPath));
                    string targetFilename = Path.GetFileNameWithoutExtension(targetPath) + ".csv";
                    JsonToCsv.Process(targetPath, targetFilename, settingFile);
                    needPressKey = false;
                }
                else
                {
                    Console.WriteLine("                     --------  {0} turn to json --------", Path.GetFileNameWithoutExtension(targetPath));
                    string targetFilename = Path.GetFileNameWithoutExtension(targetPath) + ".json";
                    needPressKey = !ExcelToJson.Process(app, targetPath, targetFilename, allSheet, needDataType);

                }
            }
            else
            {
                if (srcPath == "")
                {
                    srcPath = Directory.GetCurrentDirectory() + "\\";
                }
                
                string[] srcfiles = Directory.GetFiles(srcPath, "*.xlsx");

                for (int i = 0; i < srcfiles.Length; i++)
                {
                    string srcFilename = srcfiles[i];
                    string prename = Path.GetFileNameWithoutExtension(srcFilename).Substring(0, 2);
                    if (prename == "~$")
                    {
                        continue;
                    }
                    Console.WriteLine("                     --------  {0} --------", Path.GetFileNameWithoutExtension(srcFilename));
                    string targetFilename = targetPath + Path.GetFileNameWithoutExtension(srcFilename) + ".json";
                    bool rt = ExcelToJson.Process(app, srcFilename, targetFilename, allSheet, needDataType);
                    needPressKey = rt && needPressKey;
                }
                needPressKey = !needPressKey;
            }

            if (needPressKey)
            {
                Console.Write("press any key to quit:");
                Console.ReadKey(true);
            }
            app.Quit();
        }

        static void Main(string[] args)
        {
            Excel.Application app = new Excel.Application();
            bool needPressKey = true;
            string srcPath = "", destPath = ".";
            bool allSheet = true, needDataType = true;

            List<string> argsList = new List<string>(args);

            allSheet = GetTag(argsList, "allSheet");
            needDataType = GetTag(argsList, "needDataType");

            switch (argsList.Count)
            {
                case 1:
                    {
                        srcPath = args[0];
                        break;
                    }
                case 2:
                    {
                        srcPath = GetFileAndDirectory(args[0]);
                        destPath = GetFileAndDirectory(args[1]);
                        break;
                    }
                default:
                    {
                        Console.WriteLine("1, excel2json excelFile");
                        Console.WriteLine("2, excel2json outDir allSheet needDataType");
                        Console.WriteLine("3, excel2json srcDir outDir allSheet needDataType");
                        Console.WriteLine("4, excel2json jsonFile");

                        return;
                    }
            }
            srcPath = Path.GetFullPath(srcPath);
            destPath = Path.GetFullPath(destPath);
            if (File.Exists(srcPath))
            {
                if (Path.GetExtension(destPath) == ".json")
                {
                    needPressKey = !ExcelToJson.Process(app, srcPath, destPath, allSheet, needDataType);
                }
                else
                {
                    destPath = Path.Combine(destPath, Path.GetFileName(srcPath));
                    needPressKey = !ExcelToJson.Process(app, srcPath, destPath, allSheet, needDataType);
                }
            }
            else
            {
                if (Directory.Exists(srcPath) && Directory.Exists(destPath))
                {
                    string[] srcfiles = Directory.GetFiles(srcPath, "*.xlsx");
                    for (int i = 0; i < srcfiles.Length; i++)
                    {
                        string srcFilename = srcfiles[i];
                        string prename = Path.GetFileNameWithoutExtension(srcFilename).Substring(0, 2);
                        if (prename == "~$")
                        {
                            continue;
                        }
                        Console.WriteLine("                     --------  {0} --------", Path.GetFileNameWithoutExtension(srcFilename));
                        string targetFilename = destPath + Path.GetFileNameWithoutExtension(srcFilename) + ".json";
                        bool rt = ExcelToJson.Process(app, srcFilename, targetFilename, allSheet, needDataType);
                        needPressKey = rt && needPressKey;
                    }
                    needPressKey = !needPressKey;
                }
            }

            if (needPressKey)
            {
                Console.Write("press any key to quit:");
                Console.ReadKey(true);
            }
            app.Quit();
        }

        static string GetFileAndDirectory(string path)
        {
            if (Directory.Exists(Path.GetDirectoryName(path)) || File.Exists(path))
            {
                return path;
            }
            return "";
        }

        static bool GetTag(List<string> argsList, string tag)
        {
            for (int i = 0; i < argsList.Count; i++)
            {
                if (argsList[i].ToLower() == tag.ToLower())
                {
                    argsList.RemoveAt(i);
                    return true;
                }
            }
            return false;
        }
    }

}
