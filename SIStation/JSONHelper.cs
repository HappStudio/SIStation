using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace SIStation  
{  
    /// <summary>  
    /// Json序列化和反序列化辅助类   
    /// </summary>  
    public class JSONHelper  
    {  
        /// <summary>  
        /// Json序列化   
        /// </summary>    
        /// <param name="obj">json对象实例</param>  
        /// <returns>json字符串</returns>  
        public static string JsonSerializer(JObject obj)
        {  
            string jsonString;  
            try  
            {
                jsonString = obj.ToString();
            }  
            catch  
            {  
                jsonString = string.Empty;  
            }  
            return jsonString;  
        }  
  
  
        /// <summary>  
        /// Json反序列化  
        /// </summary>  
        /// <param name="jsonString">json字符串</param>  
        /// <returns>对象实例</returns>  
        public static JObject JsonDeserialize(string jsonString)  
        {
            JObject obj = null;
            try
            {
                obj = JObject.Parse(jsonString);
            }
            catch
            {
            }
            return obj;
        }  
  
  
        /// <summary>  
        /// 将 DataTable 序列化成 json 字符串  
        /// </summary>  
        /// <param name="dt">DataTable对象</param>  
        /// <returns>json 字符串</returns>  
        public static List<JObject> DataTableToJson(DataTable dt)  
        {  
            if (dt == null || dt.Rows.Count == 0)  
            {  
                return null;  
            }

            List<JObject> json = new List<JObject>();
            foreach (DataRow dr in dt.Rows)  
            {  
                JObject rowObj = new JObject();
                foreach (DataColumn dc in dt.Columns)  
                {  
                    rowObj.Add(dc.ColumnName, JToken.FromObject(dr[dc].ToString()));
                }
                json.Add(rowObj);
            }
            return json;  
        }

        /// <summary>  
        /// 将 json对象输出到Excel 
        /// </summary>  
        /// <param name="dt">json对象</param>  
        /// <returns>Excel</returns>  
        public static void JsonToExcel(IList<JObject> json, string excel)
        {
            Excel.Application excelApp = new Excel.Application();
            try
            {
                excelApp.DisplayAlerts = false;
                excelApp.SheetsInNewWorkbook = 1;
                excelApp.UserControl = false;
                Excel._Workbook workBook = (Excel._Workbook)(excelApp.Workbooks.Add());//添加新工作簿
                try
                {
                    workBook.Saved = true;
                    Excel._Worksheet workSheet = (Excel._Worksheet)(excelApp.Worksheets.Add());
                    workSheet.Name = excel;

                    int row = 1;
                    int column = 1;
                    foreach (JProperty property in json[0].Properties())
                    {
                        workSheet.Cells[row, column++] = property.Name;
                    }

                    foreach (JObject jobj in json)
                    {
                        column = 1;
                        row++;
                        foreach (JProperty property in jobj.Properties())
                        {
                            workSheet.Cells[row, column++] = property.Value.ToString();
                        }
                    }

                    workBook.SaveAs(Path.Combine(Directory.GetCurrentDirectory(), excel));
                }
                finally
                {
                    workBook.Close(true);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
                }
            }
            finally
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        /// <summary>  
        /// 将 Excel输出到 json对象
        /// </summary>  
        /// <param name="dt">json对象</param>  
        /// <returns>Excel</returns>  
        public static List<JObject> ExcelToJson(string excel)
        {
            List<JObject> json = new List<JObject>();
            Excel.Application excelApp = new Excel.Application();
            try
            {
                excelApp.Visible = false;
                excelApp.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), excel));
                Excel._Workbook workBook = excelApp.ActiveWorkbook;
                try
                {
                    Excel._Worksheet workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                    int column = workSheet.UsedRange.Columns.Count;
                    int row = workSheet.UsedRange.Rows.Count;

                    for (int i = 1; i < row; i++)
                    {
                        JObject rowObj = new JObject();
                        for (int j = 0; j < column; j++)
                        {
                            Object obj = workSheet.Cells[1 + i, 1 + j].Value;
                            JToken token = JToken.FromObject(obj.ToString());
                            rowObj.Add(workSheet.Cells[1, 1 + j].Value.ToString(), token);
                        }
                        json.Add(rowObj);
                    }
                }
                finally
                {
                    workBook.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
                }
            }
            finally
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            
            return json;
        }

        private static object JTokenValue(JToken token)
        {
            switch (token.Type)
            {
                case JTokenType.Integer:
                    {
                        return token.Value<int>();
                    }
                case JTokenType.Boolean:
                    {
                        return token.Value<bool>();
                    }
                case JTokenType.String:
                    {
                        return token.Value<string>();
                    }
                case JTokenType.Bytes:
                    {
                        return token.Value<byte[]>();
                    }
                default:
                    {
                        return null;
                    }
            }
        }
    }  
} 