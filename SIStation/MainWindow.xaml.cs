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
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;

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

            Excel.Application m_Excel = new Excel.Application();
            m_Excel.SheetsInNewWorkbook = 1;
            Excel._Workbook m_Book = (Excel._Workbook)(m_Excel.Workbooks.Add(Missing.Value));//添加新工作簿
            //Excel._Worksheet m_Sheet = (Excel._Worksheet)(m_Excel.Worksheets.Add(Missing.Value));
            //这里先这么写。。= =！
            Excel._Worksheet m_Sheet = (Excel._Worksheet)(m_Excel.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value));
            m_Sheet.Name = "工资报表";
            DateTime date = DateTime.Now;
            int month = date.Month;

            m_Book.SaveAs(Directory.GetCurrentDirectory()+"/MyExcel", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            m_Book.Close(false, Missing.Value, Missing.Value);
            m_Excel.Quit();
        }
    }
}
