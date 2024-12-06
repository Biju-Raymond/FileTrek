using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Drawing;

namespace FileTrek
{
    public partial class Form1 : Form
    {
        private List<FileDetails> fileDetails = new List<FileDetails>();
        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }
        
        public class FileDetails
        {
            public string FileName { get; set; } = "";
            public string FilePath { get; set; } = "";
            public DateTime LastModified { get; set; }
            public long SizeInBytes { get; set; }
        }

        private void UpdateStatus(string statusMessage, bool isSuccess = true)
        {
            string timestamp = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}: ";
            Color messageColor = isSuccess ? Color.Green : Color.Red;
            FontStyle messageFontStyle = isSuccess ? FontStyle.Regular : FontStyle.Bold;
            txtStatus.SelectionStart = txtStatus.TextLength;
            txtStatus.SelectionLength = 0;
            txtStatus.SelectionColor = Color.Black;
            txtStatus.SelectionFont = new Font(txtStatus.SelectionFont, FontStyle.Bold);
            txtStatus.AppendText(timestamp);
            txtStatus.SelectionColor = messageColor;
            txtStatus.SelectionFont = new Font(txtStatus.SelectionFont, messageFontStyle);
            txtStatus.AppendText(statusMessage + "\n");
            txtStatus.ScrollToCaret();
        }

        private async void btnUseExcelFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select an Excel file"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedExcelFile = openFileDialog.FileName;
                UpdateStatus($"Loading Excel file: {selectedExcelFile}", true);
                await ProcessFilesFromExcelAsync(selectedExcelFile);
            }
        }

        private async Task ProcessFilesFromExcelAsync(string excelFilePath)
        {
            var fileDetailsList = new List<FileDetails>();
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; 
                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++) 
                {
                    string projectName = worksheet.Cells[row, 1].Text;
                    string fileName = worksheet.Cells[row, 2].Text;
                    string filePath = worksheet.Cells[row, 3].Text;
                    if (File.Exists(filePath))
                    {
                        FileInfo file = new FileInfo(filePath);
                        fileDetailsList.Add(new FileDetails
                        {
                            FileName = fileName,
                            FilePath = filePath,
                            LastModified = file.LastWriteTime,
                            SizeInBytes = file.Length
                        });
                    }
                    await Task.Delay(10); 
                }
            }
            var sortedFiles = fileDetailsList.OrderByDescending(f => f.SizeInBytes).ToList();
            fileDetails = sortedFiles;
            UpdateStatus("Excel file processed successfully!", true);
            btnExportToExcel.Enabled = true; 
        }

        private async void btnProvidePath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderDialog.SelectedPath;
                    UpdateStatus("Initializing file processing...", true);
                    await ProcessFilesFromPathAsync(selectedPath);
                }
            }
        }

        private async Task ProcessFilesFromPathAsync(string path)
        {
            var fileDetailsList = new List<FileDetails>();
            var files = new DirectoryInfo(path).GetFiles("*", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                UpdateStatus($"Processing: {file.FullName}", true);
                fileDetailsList.Add(new FileDetails
                {
                    FileName = file.Name,
                    FilePath = file.FullName,
                    LastModified = file.LastWriteTime,
                    SizeInBytes = file.Length
                });
                await Task.Delay(10);
            }
            var sortedFiles = fileDetailsList.OrderByDescending(f => f.SizeInBytes).ToList();
            fileDetails = sortedFiles;
            UpdateStatus("File processing complete!", true);
            btnExportToExcel.Enabled = true;
        }

        private async void btnExportToExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "File_Report.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;
                await ExportToExcelAsync(fileName);
            }
        }

        private async Task ExportToExcelAsync(string fileName)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("File Report");
                worksheet.Cells[1, 1].Value = "FileName";
                worksheet.Cells[1, 2].Value = "FilePath";
                worksheet.Cells[1, 3].Value = "LastModified";
                worksheet.Cells[1, 4].Value = "SizeInBytes";
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Size = 12;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                int row = 2;
                foreach (var file in fileDetails)
                {
                    worksheet.Cells[row, 1].Value = file.FileName;
                    worksheet.Cells[row, 2].Value = file.FilePath;
                    worksheet.Cells[row, 3].Value = file.LastModified.ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cells[row, 4].Value = file.SizeInBytes;
                    row++;
                }
                long totalSize = fileDetails.Sum(f => f.SizeInBytes);
                worksheet.Cells[row, 1].Value = "Total Size";
                worksheet.Cells[row, 4].Value = totalSize;
                worksheet.Cells[row, 1, row, 4].Style.Font.Bold = true;
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                await Task.Run(() => package.SaveAs(new FileInfo(fileName)));
            }
            UpdateStatus("Report exported successfully!", true);
        }
    }
}
