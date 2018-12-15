using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel2json
{
    class ExcelToJson
    {
        public static Boolean Process(Excel.Application app, string srcFilename, string targetFilename, bool allSheet, bool needDataType)
        {

            Excel.Workbook wb = app.Workbooks.Open(srcFilename);

            //var json = new JObject();
            LitJson.JsonData exceljd = allSheet ? new LitJson.JsonData() : null;
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                //var table = new JArray() as dynamic;
                LitJson.JsonData sheetjd = new LitJson.JsonData();

                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[i + 1];
                if (ws.Name == "备注") continue;
                bool repeatId = false;

                try
                {
                    string idFormat = ws.Cells[1, 1].value.ToString();
                    if (idFormat == "[uniqueid]")
                    {
                        repeatId = false;
                    }
                    else if (idFormat == "[repeatid]")
                    {
                        repeatId = true;
                    }
                    else
                    {
                        Console.WriteLine("unknow id format sign:{0}", idFormat);
                        wb.Close();
                        return false;
                    }
                    Console.WriteLine("sheet id is: {0}", idFormat);
                }
                catch (System.Exception e)
                {
                    Console.WriteLine("sheet: {0} format error, need [uniqueid] or [repeatid].", ws.Name);
                    wb.Close();
                    return false;
                }

                sheetjd.SetJsonType(LitJson.JsonType.Object);

                Console.WriteLine("Start process sheet: {0}", ws.Name);
                for (int j = 4; j <= ws.UsedRange.Rows.Count; j++)
                {
                    Console.WriteLine("Start process row: {0}", j);
                    LitJson.JsonData rowjd = new LitJson.JsonData();
                    for (int k = 1; k <= ws.UsedRange.Columns.Count; k++)
                    {
                        string dataType = "s";
                        if (needDataType)
                        {
                            if (k == 1)
                            {
                                dataType = "i";
                            }
                            else
                            {
                                if (ws.Cells[1, k].value != null)
                                {
                                    if (ws.Cells[1, k].value != "")
                                    {
                                        dataType = ws.Cells[1, k].value;
                                    }
                                }
                            }
                        }
                        string ks, vs;
                        try
                        {
                            if (ws.Cells[3, k].value == null)
                            {
                                // 标题是空的，那么忽略这列
                                continue;
                            }
                            ks = ws.Cells[3, k].value.ToString();
                        }
                        catch (System.Exception e)
                        {
                            Console.WriteLine("error in {0} sheet: (row: {1}, col: {2}), {3}", ws.Name, 3, k, e.Message);
                            wb.Close();
                            return false;
                        }
                        if (ks == "")
                        {
                            // 标题是空的，那么忽略这列
                            continue;
                        }
                        try
                        {
                            vs = "";
                            if (ws.Cells[j, k].value != null)
                            {
                                vs = ws.Cells[j, k].value.ToString();
                            }
                        }
                        catch (System.Exception e)
                        {
                            Console.WriteLine("error in {0} sheet: (row:{1}, col: {2}), {3}", ws.Name, j, k, e.Message);
                            wb.Close();
                            return false;
                        }
                        if (needDataType)
                        {
                            try
                            {
                                if (dataType == "i")
                                {
                                    rowjd[ks] = int.Parse(vs);
                                }
                                else if (dataType == "f")
                                {
                                    rowjd[ks] = float.Parse(vs);
                                }
                                else
                                {
                                    rowjd[ks] = vs;
                                }
                            }
                            catch (System.Exception e)
                            {
                                Console.WriteLine("dataType error in {0} sheet: (row:{1}, col: {2}), {3}", ws.Name, j, k, e.Message);
                                wb.Close();
                                return false;
                            }
                        }
                        else
                        {
                            rowjd[ks] = vs;
                        }


                    }
                    try
                    {
                        string key = ws.Cells[j, 1].value.ToString();
                        if (repeatId)
                        {
                            if (!sheetjd.Keys.Contains(key))
                            {
                                sheetjd[key] = new LitJson.JsonData();
                            }

                            sheetjd[key].Add(rowjd);
                        }
                        else
                        {
                            sheetjd[key] = rowjd;
                        }

                    }
                    catch (System.Exception e)
                    {
                        Console.WriteLine("error in {0} sheet: ( row: {1}, col: {2}), {3}", ws.Name, j, 1, e.Message);
                        wb.Close();
                        return false;
                    }
                }
                if (allSheet)
                {
                    if (ws.Name == "Sheet1")
                    {
                        // 如果第一个表格没有名字，那么就取第一个表格
                        exceljd = sheetjd;
                        break;
                    }

                    exceljd[ws.Name] = sheetjd;
                }
                else
                {
                    exceljd = sheetjd;
                    // 只取第一个表
                    break;
                }
            }
            wb.Close();
            System.IO.File.WriteAllText(targetFilename, exceljd.ToJson());

            return true;
        }
    }
}
