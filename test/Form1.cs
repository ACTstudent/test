using System;
using Spire.Xls;
using System.Windows.Forms;
using System.IO;  // For File.Exists
using System.Drawing;  // For Image and PictureBox

using System.Linq;
using System.Drawing;


namespace test
{
    public partial class Form1 : Form
    {

        private Form2 form2;
        private Workbook book;
        
        

        public Form1(Form2 existingForm2)
        {
            InitializeComponent();
            this.form2 = existingForm2;

        }

        public void Update(string name, string gender, string hobbies, string favColor, string saying,
     string username, string password, string address, string email, int age, DateTime birthday,
     string course, string imagePath, int rowIndex)
        {
            txtName.Text = name;
            radMale.Checked = (gender == "Male");
            radFemale.Checked = (gender == "Female");
            chkBadminton.Checked = hobbies.Contains("Badminton");
            chkVolleyball.Checked = hobbies.Contains("Volleyball");
            chkBasketball.Checked = hobbies.Contains("Basketball");
            cbxFav.SelectedItem = favColor;
            txtSaying.Text = saying;
            txtUsername.Text = username;
            txtPassword.Text = password;
            txtAddress.Text = address;
            txtEmail.Text = email;
            numAge.Value = age;
            dtpBday.Value = birthday;
            cbxCourse.SelectedItem = course;
            txtPath.Text = imagePath;

            lblID.Text = rowIndex.ToString();  // ✅ SET lblID for later update use
        }

        public string checkEmpty()
        {


            string errors = "";

            foreach (Control C in Controls)
            {
                if (C is TextBox && string.IsNullOrEmpty(C.Text))
                {
                    errors += $"{C.Name} is empty.\n";
                }

                if (C is ComboBox comboBox && comboBox.SelectedIndex == -1)
                {
                    errors += $"{C.Name} is not selected.\n";
                }

                if (C is RadioButton radioButton)
                {
                    if (radioButton.Name == "radMale" || radioButton.Name == "radFemale")
                    {
                        if (!radMale.Checked && !radFemale.Checked)
                        {
                            errors += "Gender is not selected.\n";
                        }
                    }
                }

                if (C is CheckBox checkBox)
                {
                    if ((checkBox.Name == "chkVolleyball" || checkBox.Name == "chkBadminton" || checkBox.Name == "chkBasketball") && !chkVolleyball.Checked && !chkBadminton.Checked && !chkBasketball.Checked)
                    {
                        errors += "At least one hobby must be selected (Volleyball, Badminton, or Basketball).\n";
                    }

                    if (checkBox.Name == "chkTerms" && !checkBox.Checked)
                    {
                        errors += "Terms & Conditions are not agreed.\n";
                    }
                }
            }

            // Age validation
            if (numAge.Value <= 0)
            {
                errors += "Age must be a valid number.\n";
            }

            return errors;




        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            lblMessage.Visible = true;

            try
            {
                // Get the validation errors
                string validationErrors = checkEmpty();

                // If there are any validation errors, display them in lblMessage
                if (!string.IsNullOrEmpty(validationErrors))
                {
                    lblMessage.Text = validationErrors;  // Display errors in the label
                    lblMessage.ForeColor = Color.Red;  // You can set the color to red for errors
                    return;  // Stop further processing if there are validation errors
                }
                else
                {
                    lblMessage.Text = "";  // Clear any previous error messages
                }

                // If validation passed, proceed with collecting data
                string name = txtName.Text;
                string username = txtUsername.Text;
                string password = txtPassword.Text;
                string address = txtAddress.Text;
                string email = txtEmail.Text;
                int age = (int)numAge.Value;
                string birthday = dtpBday.Text;
                string favColor = cbxFav.Text;
                string saying = txtSaying.Text;
                string course = cbxCourse.Text;

                string gender = radMale.Checked ? "Male" : "Female";

                string hobbies = "";
                if (chkVolleyball.Checked) hobbies += chkVolleyball.Text + " ";
                if (chkBadminton.Checked) hobbies += chkBadminton.Text + " ";
                if (chkBasketball.Checked) hobbies += chkBasketball.Text + " ";

                int status = 1;

                // EXCEL
                book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");

                Worksheet sheet = book.Worksheets[0];
                int row = sheet.Rows.Length + 1;

                // Insert data into the sheet
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
                sheet.Range[row, 13].Value = status.ToString();
                sheet.Range[row, 14].Value = txtPath.Text;

                // Save the workbook back to the file
                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx", ExcelVersion.Version2016);

                // Clear all fields
                txtName.Clear();
                txtUsername.Clear();
                txtPassword.Clear();
                txtAddress.Clear();
                txtEmail.Clear();
                numAge.Value = 0;
                dtpBday.Value = DateTime.Today;
                radMale.Checked = false;
                radFemale.Checked = false;
                chkVolleyball.Checked = false;
                chkBadminton.Checked = false;
                chkBasketball.Checked = false;
                cbxFav.SelectedIndex = -1;
                txtSaying.Clear();
                cbxCourse.SelectedIndex = -1;

                form2.LoadFileFromExcel();
                this.Close();
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;  // Display error message in lblMessage
                lblMessage.ForeColor = Color.Red;  // Display the exception message in red
            }
        }






        private void btnSubmit_Click_1(object sender, EventArgs e)
        {


            int r = Convert.ToInt32(lblID.Text);

            // Submit the updated data
            string name = txtName.Text;
            string username = txtUsername.Text;
            string password = txtPassword.Text;
            string address = txtAddress.Text;
            string email = txtEmail.Text;
            int age = (int)numAge.Value;
            string birthday = dtpBday.Text;
            string favColor = cbxFav.Text;
            string saying = txtSaying.Text;
            string course = cbxCourse.Text;

            string gender = radMale.Checked ? "Male" : "Female";

            string hobbies = "";
            if (chkVolleyball.Checked) hobbies += chkVolleyball.Text + " ";
            if (chkBadminton.Checked) hobbies += chkBadminton.Text + " ";
            if (chkBasketball.Checked) hobbies += chkBasketball.Text + " ";

            // Now use the 'r' variable for updating the row in the grid
            form2.UpdatedRowInGrid(r, name, gender, hobbies, favColor, saying, username, password, address, email, age, birthday, course);

            this.Hide();


        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();

            d.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*";
            d.FilterIndex = 1;  

            // Show the file dialog
            if (d.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = d.FileName;
            }

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();

            d.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*";
            d.FilterIndex = 1;

            // Show the file dialog
            if (d.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = d.FileName;
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {




            lblMessage.Visible = true;

            try
            {
                string validationErrors = checkEmpty();
                if (!string.IsNullOrEmpty(validationErrors))
                {
                    lblMessage.Text = validationErrors;
                    lblMessage.ForeColor = Color.Red;
                    return;
                }

                // Read values
                string name = txtName.Text;
                string username = txtUsername.Text;
                string password = txtPassword.Text;
                string address = txtAddress.Text;
                string email = txtEmail.Text;
                int age = (int)numAge.Value;
                string birthday = dtpBday.Text;
                string favColor = cbxFav.Text;
                string saying = txtSaying.Text;
                string course = cbxCourse.Text;
                string gender = radMale.Checked ? "Male" : "Female";
                string hobbies = "";
                if (chkVolleyball.Checked) hobbies += chkVolleyball.Text + " ";
                if (chkBadminton.Checked) hobbies += chkBadminton.Text + " ";
                if (chkBasketball.Checked) hobbies += chkBasketball.Text + " ";

                int status = 1;

                if (!int.TryParse(lblID.Text, out int rowIndex))
                {
                    MessageBox.Show("Invalid row index for update.");
                    return;
                }

                // Update Excel file
                book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");

                Worksheet sheet = book.Worksheets[0];
                int excelRow = rowIndex + 2; // +2 because Excel is 1-based and assumes row 1 is headers

                sheet.Range[excelRow, 1].Value = name;
                sheet.Range[excelRow, 2].Value = gender;
                sheet.Range[excelRow, 3].Value = hobbies;
                sheet.Range[excelRow, 4].Value = favColor;
                sheet.Range[excelRow, 5].Value = saying;
                sheet.Range[excelRow, 6].Value = username;
                sheet.Range[excelRow, 7].Value = password;
                sheet.Range[excelRow, 8].Value = address;
                sheet.Range[excelRow, 9].Value = email;
                sheet.Range[excelRow, 10].Value = age.ToString();
                sheet.Range[excelRow, 11].Value = birthday;
                sheet.Range[excelRow, 12].Value = course;
                sheet.Range[excelRow, 13].Value = status.ToString();
                sheet.Range[excelRow, 14].Value = txtPath.Text;

                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx", ExcelVersion.Version2016);

                // Update grid
                form2.UpdatedRowInGrid(rowIndex, name, gender, hobbies, favColor, saying, username, password, address, email, age, birthday, course);

                this.Close();
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
                lblMessage.ForeColor = Color.Red;
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            lblMessage.Visible = true;

            try
            {
                // Get the validation errors
                string validationErrors = checkEmpty();

                // If there are any validation errors, display them in lblMessage
                if (!string.IsNullOrEmpty(validationErrors))
                {
                    lblMessage.Text = validationErrors;  // Display errors in the label
                    lblMessage.ForeColor = Color.Red;  // You can set the color to red for errors
                    return;  // Stop further processing if there are validation errors
                }
                else
                {
                    lblMessage.Text = "";  // Clear any previous error messages
                }

                // If validation passed, proceed with collecting data
                string name = txtName.Text;
                string username = txtUsername.Text;
                string password = txtPassword.Text;
                string address = txtAddress.Text;
                string email = txtEmail.Text;
                int age = (int)numAge.Value;
                string birthday = dtpBday.Text;
                string favColor = cbxFav.Text;
                string saying = txtSaying.Text;
                string course = cbxCourse.Text;

                string gender = radMale.Checked ? "Male" : "Female";

                string hobbies = "";
                if (chkVolleyball.Checked) hobbies += chkVolleyball.Text + " ";
                if (chkBadminton.Checked) hobbies += chkBadminton.Text + " ";
                if (chkBasketball.Checked) hobbies += chkBasketball.Text + " ";

                int status = 1;

                // EXCEL
                book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");

                Worksheet sheet = book.Worksheets[0];
                int row = sheet.Rows.Length + 1;

                // Insert data into the sheet
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
                sheet.Range[row, 13].Value = status.ToString();
                sheet.Range[row, 14].Value = txtPath.Text;

                // Save the workbook back to the file
                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx", ExcelVersion.Version2016);

                // Clear all fields
                txtName.Clear();
                txtUsername.Clear();
                txtPassword.Clear();
                txtAddress.Clear();
                txtEmail.Clear();
                numAge.Value = 0;
                dtpBday.Value = DateTime.Today;
                radMale.Checked = false;
                radFemale.Checked = false;
                chkVolleyball.Checked = false;
                chkBadminton.Checked = false;
                chkBasketball.Checked = false;
                cbxFav.SelectedIndex = -1;
                txtSaying.Clear();
                cbxCourse.SelectedIndex = -1;

                form2.LoadFileFromExcel();
                this.Close();
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;  // Display error message in lblMessage
                lblMessage.ForeColor = Color.Red;  // Display the exception message in red
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void rad_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dtpBday_ValueChanged(object sender, EventArgs e)
        {
            DateTime birthDate = dtpBday.Value;
            int age = DateTime.Today.Year - birthDate.Year;

            if (birthDate.Date > DateTime.Today.AddYears(-age)) age--;

            numAge.Value = age;
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}