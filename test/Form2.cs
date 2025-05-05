using Spire.Xls;
using System;
using System.Data;
using System.Net;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using System.Security.Cryptography;
using System.Linq;

namespace test
{

    public partial class Form2 : Form
    {
     

        public Form2()
        {
            InitializeComponent();
            LoadFileFromExcel();

            
        }


        public void LoadFileFromExcel()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];

            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();
            dtgDisplay.DataSource = dt;
        }


        private void btnClose_Click(object sender, EventArgs e)
        {
        
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
            MessageBox.Show("Successfully deleted.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //dtgDisplay.Columns["Status"].Visible = false;

        }


        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dtgDisplay.CurrentRow != null && dtgDisplay.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dtgDisplay.SelectedRows[0].Index;

                // Ensure the selected index is valid before proceeding
                if (selectedRowIndex >= 0 && selectedRowIndex < dtgDisplay.Rows.Count)
                {
                    // Call your required methods here before removing the row
                    Sheets(0, 1, selectedRowIndex);
                    LoadFileFromExcel();

                    // Make sure the Panel_Inactive form is available
                    Panel_Inactive inact = Application.OpenForms.OfType<Panel_Inactive>().FirstOrDefault();
                    if (inact != null)
                    {
                        inact.LoadFileFromExcelInactive();

                        // Log the action
                        MyLogs logs = new MyLogs();
                        logs.InsertLogs(Event.GetUser, "Transferred an active student to the inactive list.");
                    }

                    // Remove the row after the necessary operations
                    // Always check if the row is still there before removing it
                    if (selectedRowIndex >= 0 && selectedRowIndex < dtgDisplay.Rows.Count)
                    {
                        dtgDisplay.Rows.RemoveAt(selectedRowIndex);

                        // Log the delete operation
                        MyLogs log = new MyLogs();
                        log.InsertLogs(Event.GetUser, "Successfully deleted.");
                        MessageBox.Show("Successfully deleted.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Row has already been removed or is invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid row index. Cannot delete the row.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string Search = txtSearch.Text; 

            
            bool foundMatch = false; 

            foreach (DataGridViewRow row in dtgDisplay.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool searchMatch = false;

                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null && cell.Value.ToString().ToLower().Contains(Search.ToLower()))
                        {
                            searchMatch = true;
                            foundMatch = true; 
                            break;
                        }
                    }
                    row.Visible = searchMatch;
                }

            //seacrh not match
            if (!foundMatch)
                {
                    MessageBox.Show("No matching records found.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
        }



        private void dtgDisplay_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            if (dtgDisplay.SelectedRows.Count > 0)
            {
                int rowIndex = dtgDisplay.SelectedRows[0].Index;

                // Retrieve all data from the selected row
                string name = dtgDisplay.Rows[rowIndex].Cells[0].Value.ToString();
                string gender = dtgDisplay.Rows[rowIndex].Cells[1].Value.ToString();
                string hobbies = dtgDisplay.Rows[rowIndex].Cells[2].Value.ToString();
                string favColor = dtgDisplay.Rows[rowIndex].Cells[3].Value.ToString();
                string saying = dtgDisplay.Rows[rowIndex].Cells[4].Value.ToString();
                string username = dtgDisplay.Rows[rowIndex].Cells[5].Value.ToString();
                string password = dtgDisplay.Rows[rowIndex].Cells[6].Value.ToString();
                string address = dtgDisplay.Rows[rowIndex].Cells[7].Value.ToString();
                string email = dtgDisplay.Rows[rowIndex].Cells[8].Value.ToString();
                int age = int.Parse(dtgDisplay.Rows[rowIndex].Cells[9].Value.ToString());
                DateTime birthday = DateTime.Parse(dtgDisplay.Rows[rowIndex].Cells[10].Value.ToString());
                string course = dtgDisplay.Rows[rowIndex].Cells[11].Value.ToString();

                // Set the DateTimePicker's value
                // Retrieve the profile picture path (stored in column 14)
                string imagePath = dtgDisplay.Rows[rowIndex].Cells[13].Value.ToString();

                // Create a new instance of Form1 and pass data including the profile picture path (as text)
                Form1 form1 = new Form1(this);
                form1.Update(name, gender, hobbies, favColor, saying, username, password, address, email, age, birthday, course, imagePath, rowIndex);
                form1.Show();
            }
            else
            {
                MessageBox.Show("Please select a row to update.");
            }
        }

        // update the selected row in the DataGridView
        public void UpdatedRowInGrid( int r, string name, string gender, string hobbies, string favColor, string saying, string username, string password, string address, string email, int age, string birthday, string course)
        {
            
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
                Worksheet sheet = book.Worksheets[0];

                int row = r + 2;
                sheet.Range[row, 1].Value = name;
                sheet.Range[row, 2].Value = gender;
                sheet.Range[row, 3].Value = hobbies;
                sheet.Range[row, 4].Value = favColor;
                sheet.Range[row, 5].Value = saying;
                sheet.Range[row, 6].Value = username;
                sheet.Range[row, 7].Value = password;
                sheet.Range[row, 8].Value = address;
                sheet.Range[row, 9].Value = email;
                sheet.Range[row, 10].Value = age.ToString();
                sheet.Range[row, 11].Value = birthday;
                sheet.Range[row, 12].Value = course;

                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx", ExcelVersion.Version2016);


                int dtgIndex = r;
                dtgDisplay.Rows[dtgIndex].Cells[0].Value = name;
                dtgDisplay.Rows[dtgIndex].Cells[1].Value = gender;
                dtgDisplay.Rows[dtgIndex].Cells[2].Value = hobbies;
                dtgDisplay.Rows[dtgIndex].Cells[3].Value = favColor;
                dtgDisplay.Rows[dtgIndex].Cells[4].Value = saying;
                dtgDisplay.Rows[dtgIndex].Cells[5].Value = username;
                dtgDisplay.Rows[dtgIndex].Cells[6].Value = password;
                dtgDisplay.Rows[dtgIndex].Cells[7].Value = address;
                dtgDisplay.Rows[dtgIndex].Cells[8].Value = email;
                dtgDisplay.Rows[dtgIndex].Cells[9].Value = age;
                dtgDisplay.Rows[dtgIndex].Cells[10].Value = birthday;
                dtgDisplay.Rows[dtgIndex].Cells[11].Value = course;


            
        }
    

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1(this);
            form.Show();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadFileFromExcel();

            

        }

        private void dtgDisplay_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1(this);
            form.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            LoadFileFromExcel();

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            string Search = txtSearch.Text;


            bool foundMatch = false;

            foreach (DataGridViewRow row in dtgDisplay.Rows)
            {
                if (row.IsNewRow) continue;

                bool searchMatch = false;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToLower().Contains(Search.ToLower()))
                    {
                        searchMatch = true;
                        foundMatch = true;
                        break;
                    }
                }
                row.Visible = searchMatch;
            }

            //seacrh not match
            if (!foundMatch)
            {
                MessageBox.Show("No matching records found.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (dtgDisplay.CurrentRow != null && dtgDisplay.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dtgDisplay.SelectedRows[0].Index;

                // Ensure the selected index is valid before proceeding
                if (selectedRowIndex >= 0 && selectedRowIndex < dtgDisplay.Rows.Count)
                {
                    // Call your required methods here before removing the row
                    Sheets(0, 1, selectedRowIndex);
                    LoadFileFromExcel();

                    // Make sure the Panel_Inactive form is available
                    Panel_Inactive inact = Application.OpenForms.OfType<Panel_Inactive>().FirstOrDefault();
                    if (inact != null)
                    {
                        inact.LoadFileFromExcelInactive();

                        // Log the action
                        MyLogs logs = new MyLogs();
                        logs.InsertLogs(Event.GetUser, "Transferred an active student to the inactive list.");
                    }

                    // Remove the row after the necessary operations
                    // Always check if the row is still there before removing it
                    if (selectedRowIndex >= 0 && selectedRowIndex < dtgDisplay.Rows.Count)
                    {
                        dtgDisplay.Rows.RemoveAt(selectedRowIndex);

                        // Log the delete operation
                        MyLogs log = new MyLogs();
                        log.InsertLogs(Event.GetUser, "Successfully deleted.");
                        MessageBox.Show("Successfully deleted.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Row has already been removed or is invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid row index. Cannot delete the row.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
    }

}