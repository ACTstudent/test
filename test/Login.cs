using Spire.Xls;
using System;
using System.Windows.Forms;


namespace test
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
            lblErrorMessage.Visible = false;

        }

        private void btnLgin_Click(object sender, EventArgs e)
        {

            // Workbook Credentials
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            int row = sheet.LastRow;

            bool loginSuccess = false;


            for (int i = 2; i <= row; i++)
            {
                string user = sheet.Range[i, 6].Value;
                string pass = sheet.Range[i, 7].Value;

                if (user == txtUser.Text && pass == txtPass.Text)
                {
                    Event.GetUser = txtUser.Text;
                    Event.Name = sheet.Range[i, 1].Value;
                    Event.Path = sheet.Range[i, 14].Value;
                    loginSuccess = true;
                    break;


                }
                else
                {
                    loginSuccess = false;
                }

            }
            if (loginSuccess == true)
            {
                MessageBox.Show("Login Successfully!", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Dashboard dashboard = new Dashboard();
                this.Hide();
                dashboard.Show();

                MyLogs log = new MyLogs();
                log.InsertLogs(Event.GetUser, "Successfully logged in.");
            }
            else
            {
                MessageBox.Show("Invalid account.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }



        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void cbxShowandHide_CheckedChanged(object sender, EventArgs e)
        {
            if (cbxShowandHide.Checked)
            {
                txtPass.UseSystemPasswordChar = true;
            }
            else
            {
                txtPass.UseSystemPasswordChar = false;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

            Application.Exit();
        }

        private void label1_Click(object sender, EventArgs e)
        {
         
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            // Workbook Credentials
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            int row = sheet.LastRow;

            bool loginSuccess = false;

            for (int i = 2; i <= row; i++)
            {
                string user = sheet.Range[i, 6].Value;
                string pass = sheet.Range[i, 7].Value;

                if (user == txtUser.Text && pass == txtPass.Text)
                {
                    Event.GetUser = txtUser.Text;
                    Event.Name = sheet.Range[i, 1].Value;
                    Event.Path = sheet.Range[i, 14].Value;
                    loginSuccess = true;
                    break;
                }
            }

            if (loginSuccess)
            {
                lblErrorMessage.Visible = false; // Hide error message

                Dashboard dashboard = new Dashboard();
                this.Hide();
                dashboard.Show();

                MyLogs log = new MyLogs();
                log.InsertLogs(Event.GetUser, "Successfully logged in.");
            }
            else
            {
                lblErrorMessage.Text = "Invalid username or password.";
                lblErrorMessage.Visible = true;
            }
        }

        private void txtPass_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }
    }
}
