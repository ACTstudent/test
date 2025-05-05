using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    internal class MyLogs
    { 

            public void InsertLogs (string user, string message)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
                Worksheet sh = book.Worksheets[2];

                sh.InsertRow(2, 1);
                sh.Range[2, 1].Value = user;
                sh.Range[2, 2].Value = message;
                sh.Range[2, 3].Value = DateTime.Now.ToString("MM,dd,yyyy");
                sh.Range[2, 4].Value = DateTime.Now.ToString("hh : mm : ss : tt");

                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx", ExcelVersion.Version2016);

            }
    }
}
