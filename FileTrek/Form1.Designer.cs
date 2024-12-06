namespace FileTrek
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnProvidePath = new System.Windows.Forms.Button();
            this.btnUseExcelFile = new System.Windows.Forms.Button();
            this.btnExportToExcel = new System.Windows.Forms.Button();
            this.txtStatus = new System.Windows.Forms.RichTextBox();
            this.btnProvidePath.Location = new System.Drawing.Point(50, 30);
            this.btnProvidePath.Name = "btnProvidePath";
            this.btnProvidePath.Size = new System.Drawing.Size(150, 50);
            this.btnProvidePath.TabIndex = 1;
            this.btnProvidePath.Text = "Provide Path";
            this.btnProvidePath.UseVisualStyleBackColor = true;
            this.btnProvidePath.Click += new System.EventHandler(this.btnProvidePath_Click);
            this.btnUseExcelFile.Location = new System.Drawing.Point(220, 30);
            this.btnUseExcelFile.Name = "btnUseExcelFile";
            this.btnUseExcelFile.Size = new System.Drawing.Size(150, 50);
            this.btnUseExcelFile.TabIndex = 2;
            this.btnUseExcelFile.Text = "Use Excel File";
            this.btnUseExcelFile.UseVisualStyleBackColor = true;
            this.btnUseExcelFile.Click += new System.EventHandler(this.btnUseExcelFile_Click);
            this.btnExportToExcel.Location = new System.Drawing.Point(390, 30);
            this.btnExportToExcel.Name = "btnExportToExcel";
            this.btnExportToExcel.Size = new System.Drawing.Size(150, 50);
            this.btnExportToExcel.TabIndex = 3;
            this.btnExportToExcel.Text = "Export to Excel";
            this.btnExportToExcel.UseVisualStyleBackColor = true;
            this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);
            this.btnExportToExcel.Enabled = false;  
            this.txtStatus.Location = new System.Drawing.Point(50, 120);
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ReadOnly = true;
            this.txtStatus.Size = new System.Drawing.Size(700, 250);
            this.txtStatus.TabIndex = 4;
            this.txtStatus.Text = "";
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnProvidePath);
            this.Controls.Add(this.btnUseExcelFile);
            this.Controls.Add(this.btnExportToExcel);
            this.Controls.Add(this.txtStatus);
            this.Name = "Form1";
            this.Text = "FileTrek";
        }
        private System.Windows.Forms.Button btnProvidePath;
        private System.Windows.Forms.Button btnUseExcelFile;
        private System.Windows.Forms.Button btnExportToExcel;
        private System.Windows.Forms.RichTextBox txtStatus;
    }
}
