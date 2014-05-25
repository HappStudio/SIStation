using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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

            List<JObject> tableData = new List<JObject>();
            foreach (DataRow dr in dt.Rows)  
            {  
                JObject rowObj = new JObject();
                foreach (DataColumn dc in dt.Columns)  
                {  
                    rowObj.Add(dc.ColumnName, JToken.FromObject(dr[dc]));
                }
                tableData.Add(rowObj);
            }
            return tableData;  
        }

        /// <summary>  
        /// 将 json对象输出到Excel 
        /// </summary>  
        /// <param name="dt">json对象</param>  
        /// <returns>Excel</returns>  
        
    }  
} 