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
    public partial class Panel_Inactive : Form
    {
        public Panel_Inactive()
        {
            InitializeComponent();
            LoadFileFromExcelInactive();
        }
        public void LoadFileFromExcelInactive()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
            Worksheet sheet = book.Worksheets[1];

            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();
            dtgDisplayInact.DataSource = dt;
        }

        public void Sheets(int fromSheetIndex, int toSheetIndex, int rowIndexToMove)
        {

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");

            Worksheet fromSheet = book.Worksheets[fromSheetIndex];
            Worksheet toSheet = book.Worksheets[toSheetIndex];

            int lastRow = toSheet.LastRow + 1;

            for (int col = 1; col <= fromSheet.LastColumn; col++)
            {
                var value = fromSheet.Range[rowIndexToMove + 2, col].Value ?? string.Empty;


                if (col == 13)
                {
                    value = (toSheetIndex == 0) ? "1" : "0";

                }
                toSheet.Range[lastRow, col].Text = value.ToString();
            }


            fromSheet.DeleteRow(rowIndexToMove + 2);

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx", ExcelVersion.Version2016);

            MessageBox.Show("The student has been successfully transferred!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dtgDisplayInact.Columns["Status"].Visible = false;
        }

        private void btnActive_Click(object sender, EventArgs e)
        {
            MyLogs logs = new MyLogs();
            logs.InsertLogs(Event.GetUser, "Transferred a inactive Student in the active list.");
            if (dtgDisplayInact.CurrentRow != null)
            {

                int selectedRowIndex = dtgDisplayInact.CurrentRow.Index;


                Sheets(1, 0, selectedRowIndex);

                LoadFileFromExcelInactive();
                Form2 form2 = Application.OpenForms.OfType<Form2>().FirstOrDefault();
                if (form2 != null)
                {
                    form2.LoadFileFromExcel();
                }
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadFileFromExcelInactive();
        }

        private void Panel_Inactive_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            MyLogs logs = new MyLogs();
            logs.InsertLogs(Event.GetUser, "Transferred a inactive Student in the active list.");
            if (dtgDisplayInact.CurrentRow != null)
            {

                int selectedRowIndex = dtgDisplayInact.CurrentRow.Index;


                Sheets(1, 0, selectedRowIndex);

                LoadFileFromExcelInactive();
                Form2 form2 = Application.OpenForms.OfType<Form2>().FirstOrDefault();
                if (form2 != null)
                {
                    form2.LoadFileFromExcel();
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            LoadFileFromExcelInactive();

        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            MyLogs logs = new MyLogs();
            logs.InsertLogs(Event.GetUser, "Transferred a inactive Student in the active list.");
            if (dtgDisplayInact.CurrentRow != null)
            {

                int selectedRowIndex = dtgDisplayInact.CurrentRow.Index;


                Sheets(1, 0, selectedRowIndex);

                LoadFileFromExcelInactive();
                Form2 form2 = Application.OpenForms.OfType<Form2>().FirstOrDefault();
                if (form2 != null)
                {
                    form2.LoadFileFromExcel();
                }
            }
        }

        private void dtgDisplayInact_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
