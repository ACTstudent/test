using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    public partial class Panel_Logs : Form
    {
        public Panel_Logs()
        {
            InitializeComponent();
            LoadFileFromExcelLogs();
        }

        public void LoadFileFromExcelLogs()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
            Worksheet sheet = book.Worksheets[2];

            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();
            dtgLogs.DataSource = dt;
        }

        private void Panel_Logs_Load(object sender, EventArgs e)
        {

        }

        private void dtgLogs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
