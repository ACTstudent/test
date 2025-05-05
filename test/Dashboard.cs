using Spire.Xls;
using System;
using System.IO;  // For File.Exists
using System.Drawing;  // For Image and PictureBox

using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Windows.Forms;

namespace test
{
    public partial class Dashboard : Form
    {
        private Form2 form2;
        Workbook book = new Workbook();

        public Dashboard()
        {
            InitializeComponent();
            form2 = new Form2();
            form2.Owner = this;
            LoadCount();

            if (!string.IsNullOrEmpty(Event.Path) && File.Exists(Event.Path))
            {
                pcbProfile.Image = Image.FromFile(Event.Path);
            }
            else
            {
                return;
            }
        }




        public void LoadCount()
        {

            // Load Excel workbook
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\ARDIMER\Book.xlsx");
            Worksheet activeSheet = book.Worksheets[0];
            Worksheet inactiveSheet = book.Worksheets[1];

            // Initialize counters
            int male = 0, female = 0;
            int black = 0, white = 0, gray = 0;
            int volleyball = 0, badminton = 0, basketball = 0;
            int BSIT = 0, BSCS = 0, BSCPE = 0;

            // Method to process each sheet
            void sheetCount(Worksheet sheet)
            {
                for (int i = 2; i <= sheet.LastRow; i++) // Assuming row 1 is the header
                {
                    
                        string gender = sheet.Range[i, 2]?.Text?.Trim() ?? "";
                        if (gender == "Male") male++;
                        else if (gender == "Female") female++;

                        string hobby = sheet.Range[i, 3]?.Text?.Trim() ?? "";
                        if (hobby == "Volleyball") volleyball++;
                        else if (hobby == "Basketball") basketball++;
                        else if (hobby == "Badminton") badminton++;

                        string color = sheet.Range[i, 4]?.Text?.Trim() ?? "";
                        if (color == "Black") black++;
                        else if (color == "White") white++;
                        else if (color == "Gray") gray++;

                        string course = sheet.Range[i, 12]?.Text?.Trim() ?? "";
                        if (course == "BSIT") BSIT++;
                        else if (course == "BSCS") BSCS++;
                        else if (course == "BSCPE") BSCPE++;
                    
                }
            }

            // Process both sheets
            sheetCount(activeSheet);
            sheetCount(inactiveSheet);

            // Update UI labels
            lblActive.Text = (activeSheet.LastRow - 1).ToString();   // Assuming first row is header
            lblInactive.Text = (inactiveSheet.LastRow - 1).ToString();

            lblMale.Text = male.ToString();
            lblFemale.Text = female.ToString();

            lblBlack.Text = black.ToString();
            lblWhite.Text = white.ToString();
            lblGray.Text = gray.ToString();

            lblVolleyball.Text = volleyball.ToString();
            lblBadminton.Text = badminton.ToString();
            lblBasketball.Text = basketball.ToString();

            lblBSIT.Text = BSIT.ToString();
            lblBSCS.Text = BSCS.ToString();
            lblBSCPE.Text = BSCPE.ToString();

            lblName.Text = Event.Name;
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");



        }






        private void btnInactive_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Clicked Inactive.");

            mainPanel.Show();
            mainPanel.Controls.Clear();
            Panel_Inactive inac = new Panel_Inactive();
            inac.TopLevel = false;
            mainPanel.BringToFront();
            mainPanel.Controls.Add(inac);

            inac.Show();

            // Refresh the count data after the action
            LoadCount();  // This will update the label values on the dashboard

        }


        private void btnLogs_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Clicked logs.");

            mainPanel.Show();
            mainPanel.Controls.Clear();
            Panel_Logs logs = new Panel_Logs();
            logs.TopLevel = false;
            mainPanel.BringToFront();
            mainPanel.Controls.Add(logs);
            logs.BringToFront();

            logs.Show();


            

            
        }

        private void btnActive_Click(object sender, EventArgs e)
        {


            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Clicked Active.");

            mainPanel.Show();
            mainPanel.Controls.Clear();

            Form2 act = new Form2(); // ← your Panel_Active Form
            act.TopLevel = false;
            mainPanel.BringToFront();
            mainPanel.Controls.Add(act);

            act.Show();

            // Refresh the count data after showing the new panel
            LoadCount();  // This will update the labels with the latest counts


        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Logged out.");

            Login login = new Login();
            this.Close();
            login.Show();

            
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {
            LoadCount();
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Returned to Main Panel.");
            mainPanel.Hide();

            LoadCount();  // This will update the labels with the latest counts



        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
                
        }

        private void mainPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Logged out.");

            Login login = new Login();
            this.Close();
            login.Show();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Returned to Main Panel.");
            mainPanel.Hide();

            LoadCount();  // This will update the labels with the latest counts
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Clicked Active.");

            mainPanel.Show();
            mainPanel.Controls.Clear();

            Form2 act = new Form2(); // ← your Panel_Active Form
            act.TopLevel = false;
            mainPanel.BringToFront();
            mainPanel.Controls.Add(act);

            act.Show();

            // Refresh the count data after showing the new panel
            LoadCount();  // This will update the labels with the latest counts
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Clicked Inactive.");

            mainPanel.Show();
            mainPanel.Controls.Clear();
            Panel_Inactive inac = new Panel_Inactive();
            inac.TopLevel = false;
            mainPanel.BringToFront();
            mainPanel.Controls.Add(inac);

            inac.Show();

            // Refresh the count data after the action
            LoadCount();  // This will update the label values on the dashboard
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            MyLogs log = new MyLogs();
            log.InsertLogs(Event.GetUser, "Clicked logs.");

            mainPanel.Show();
            mainPanel.Controls.Clear();
            Panel_Logs logs = new Panel_Logs();
            logs.TopLevel = false;
            mainPanel.BringToFront();
            mainPanel.Controls.Add(logs);
            logs.BringToFront();

            logs.Show();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void panel9_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
