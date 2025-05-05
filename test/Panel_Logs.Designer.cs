namespace test
{
    partial class Panel_Logs
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Panel_Logs));
            this.dtgLogs = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dtgLogs)).BeginInit();
            this.SuspendLayout();
            // 
            // dtgLogs
            // 
            this.dtgLogs.AllowUserToAddRows = false;
            this.dtgLogs.AllowUserToResizeColumns = false;
            this.dtgLogs.AllowUserToResizeRows = false;
            this.dtgLogs.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.dtgLogs.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dtgLogs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgLogs.Location = new System.Drawing.Point(12, -4);
            this.dtgLogs.Name = "dtgLogs";
            this.dtgLogs.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dtgLogs.Size = new System.Drawing.Size(915, 406);
            this.dtgLogs.TabIndex = 1;
            this.dtgLogs.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgLogs_CellContentClick);
            // 
            // Panel_Logs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(928, 404);
            this.Controls.Add(this.dtgLogs);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Panel_Logs";
            this.Text = "Panel_Logs";
            this.Load += new System.EventHandler(this.Panel_Logs_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtgLogs)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dtgLogs;
    }
}