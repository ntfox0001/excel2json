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
            string settingFile = System.IO.Directory.GetCurrentDirectory() + "\\columnSetting.xml";
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

            if (System.IO.File.Exists(targetPath))
            {
                if (System.IO.Path.GetExtension(targetPath) == ".json")
                {
                    Console.WriteLine("                     --------  {0} turn to csv --------", System.IO.Path.GetFileNameWithoutExtension(targetPath));
                    string targetFilename = System.IO.Path.GetFileNameWithoutExtension(targetPath) + ".csv";
                    JsonToCsv.Process(targetPath, targetFilename, settingFile);
                    needPressKey = false;
                }
                else
                {
                    Console.WriteLine("                     --------  {0} turn to json --------", System.IO.Path.GetFileNameWithoutExtension(targetPath));
                    string targetFilename = System.IO.Path.GetFileNameWithoutExtension(targetPath) + ".json";
                    needPressKey = !ExcelToJson.Process(app, targetPath, targetFilename, allSheet, needDataType);

                }
            }
            else
            {
                if (srcPath == "")
                {
                    srcPath = System.IO.Directory.GetCurrentDirectory() + "\\";
                }
                
                string[] srcfiles = System.IO.Directory.GetFiles(srcPath, "*.xlsx");

                for (int i = 0; i < srcfiles.Length; i++)
                {
                    string srcFilename = srcfiles[i];
                    string prename = System.IO.Path.GetFileNameWithoutExtension(srcFilename).Substring(0, 2);
                    if (prename == "~$")
                    {
                        continue;
                    }
                    Console.WriteLine("                     --------  {0} --------", System.IO.Path.GetFileNameWithoutExtension(srcFilename));
                    string targetFilename = targetPath + System.IO.Path.GetFileNameWithoutExtension(srcFilename) + ".json";
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
            srcPath = System.IO.Path.GetFullPath(srcPath);
            destPath = System.IO.Path.GetFullPath(destPath);
            if (System.IO.File.Exists(srcPath))
            {
                if (System.IO.File.Exists(destPath))
                {
                    needPressKey = !ExcelToJson.Process(app, srcPath, destPath, allSheet, needDataType);
                }
                else
                {
                    destPath = System.IO.Path.Combine(destPath, System.IO.Path.GetFileName(srcPath));
                    needPressKey = !ExcelToJson.Process(app, srcPath, destPath, allSheet, needDataType);
                }
            }
            else
            {
                if (System.IO.Directory.Exists(srcPath) && System.IO.Directory.Exists(destPath))
                {
                    string[] srcfiles = System.IO.Directory.GetFiles(srcPath, "*.xlsx");
                    for (int i = 0; i < srcfiles.Length; i++)
                    {
                        string srcFilename = srcfiles[i];
                        string prename = System.IO.Path.GetFileNameWithoutExtension(srcFilename).Substring(0, 2);
                        if (prename == "~$")
                        {
                            continue;
                        }
                        Console.WriteLine("                     --------  {0} --------", System.IO.Path.GetFileNameWithoutExtension(srcFilename));
                        string targetFilename = destPath + System.IO.Path.GetFileNameWithoutExtension(srcFilename) + ".json";
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
            if (System.IO.Directory.Exists(Path.GetDirectoryName(path)) || System.IO.File.Exists(path))
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
