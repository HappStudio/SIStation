using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Data;
using System.Data.SQLite;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace SIStation
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            try
            {
                string sql = "create table if not exists UserTable (userId integer primary key AutoIncrement,displayName varchar(20),department varchar(20))";
                SQLiteHelper.Instance.ExecuteNonQuery(sql);

                sql = "select * from UserTable";
                DataTable td = SQLiteHelper.Instance.ExecuteReader(sql);
                if (td.Rows.Count==0)
                {
                    sql = "insert into UserTable (userId, displayName, department) values (10, '妖怪', '321')";
                    SQLiteHelper.Instance.ExecuteNonQuery(sql);
                }

                sql = "select userId, displayName, department from UserTable where userId = @userid";
                td = SQLiteHelper.Instance.ExecuteReader(sql, new SQLiteParameter("userid", 10));
                List<JObject> json = JSONHelper.DataTableToJson(td);

                Debug.Assert("妖怪".Equals(json[0].Property("displayName").Value.ToString()), "WTF!!!");

                sql = "select * from UserTable";
                td = SQLiteHelper.Instance.ExecuteReader(sql);
                json = JSONHelper.DataTableToJson(td);
                JSONHelper.JsonToExcel(json, "员工花名册");

               // List<JObject> jsonValidated = JSONHelper.ExcelToJson("员工花名册");
               // Debug.Assert(JSONHelper.JsonSerializer(jsonValidated[0]).Equals(JSONHelper.JsonSerializer(json[0])), "WTF!!!");
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.StackTrace);
            }
        }
    }
}
