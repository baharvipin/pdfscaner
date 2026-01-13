using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.IO; 
using PdfReader.BusinessLogic;
using OfficeOpenXml.DataValidation;
using PdfReader.BusinessLogic.Model;
using System.Xml.Linq;
using System.Collections;
using OfficeOpenXml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PdfReader
{
    public partial class PdfScan : Form
    {
        PdfExtractor pdfExtractor;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.DataGridView dataGridView;
        string pdfPath = ConfigurationManager.AppSettings["PdfPath"];
        string excelName = ConfigurationManager.AppSettings["ExcelName"];
        private string selectedFolderPath;
        private Dictionary<string, DataGridView> dataGridViewDictionary;
        public PdfScan()
        {
            InitializeComponent();
            dataGridViewDictionary = new Dictionary<string, DataGridView>();
            pdfExtractor = new PdfExtractor();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            string excelName = ConfigurationManager.AppSettings["ExcelName"];
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = $"{basePath}\\UploadedExcel\\{excelName}";
            if (File.Exists(destinationFolderPath))
            {
                this.uploadButton.Text = "Replace Excel";
                tabControl1.Visible = true;
            }
            else
            {
                this.uploadButton.Text = "Upload Excel";
                tabControl1.Visible = false;
            } 

            // Add the TabControl to the Form
            this.Controls.Add(tabControl1);

            // Subscribe to the SelectedIndexChanged event
            this.tabControl1.SelectedIndexChanged += TabControl_SelectedIndexChanged;
        }
         
        private void Form1_Load(object sender, EventArgs e)
        {
            // Ensure the first tab is selected
            tabControl1.SelectedIndex = 0;
            // Load data into the DataGridView on the first tab
            LoadDataIntoFirstTab();
            tabControl1.Selecting += new TabControlCancelEventHandler(tabControl1_Selecting);
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            // Example condition: Disable the tabs
            if (string.IsNullOrEmpty(selectedFolderPath) && e.TabPageIndex != 5)
            {
                // Cancel the selection of the tab, preventing the user from accessing it
                e.Cancel = true;
            }
        }
        private void tabPage1_Click(object sender, EventArgs e)
        {  

        }
         
        public void LoadDataIntoFirstTab()
        {
            loaderLabel.Visible = true; // Show the loader
            dataGridViewForSourceExcel.Visible = false;
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel");

            // Ensure the destination directory exists
            if (!Directory.Exists(destinationFolderPath))
            {
                Directory.CreateDirectory(destinationFolderPath);
            }

            string destinationFilePath = Path.Combine(destinationFolderPath, excelName);
            if (File.Exists(destinationFilePath))
            {
                sourceExcelCount.Text = "";
                dataGridViewForSourceExcel.DataSource = new DataTable();
                PdfExtractionResult result = PdfExtractor.ExtractExcelTuple();
                
                dataGridViewForSourceExcel.DataSource = result.SourceExcelTable; 
                AddOrUpdateDataGridView("sourceExcel", dataGridViewForSourceExcel);

                // Determine the row count based on the DataSource type
                UpdateRowCountLabel();

                dataGridViewForSourceExcel.Refresh();
                sourceExcelCount.Refresh();
                loaderLabel.Visible = false; // Show the loader
                dataGridViewForSourceExcel.Visible = true;
            }
        }

        private void UpdateRowCountLabelForInvoiceNotFoundInExcel()
        {

            if (invoiceButNotMatchedDt.DataSource == null)
            {
                invoiceNotFoundExcelCount.Text = "No data loaded.";
                return;
            }

            int rowCount = 0;

            if (invoiceButNotMatchedDt.DataSource is DataTable dataTable)
            {
                rowCount = dataTable.Rows.Count;
            }
            else if (invoiceButNotMatchedDt.DataSource is BindingSource bindingSource)
            {
                rowCount = bindingSource.Count;
            }
            else if (invoiceButNotMatchedDt.DataSource is IList list)
            {
                rowCount = list.Count;
            }
            else
            {
                rowCount = invoiceButNotMatchedDt.Rows.Count;
            }

            // Update label text
            invoiceNotFoundExcelCount.Text = $"Invoice Not Found Excel Count: {rowCount}";
        }

        private void UpdateRowCountLabelForInvoiceAbsentPDF()
        {

            if (dataGridViewNotMatchedInvoice.DataSource == null)
            {
                invoiceAbsentInPDFCount.Text = "No data loaded.";
                return;
            }

            int rowCount = 0;

            if (dataGridViewNotMatchedInvoice.DataSource is DataTable dataTable)
            {
                rowCount = dataTable.Rows.Count;
            }
            else if (dataGridViewNotMatchedInvoice.DataSource is BindingSource bindingSource)
            {
                rowCount = bindingSource.Count;
            }
            else if (dataGridViewNotMatchedInvoice.DataSource is IList list)
            {
                rowCount = list.Count;
            }
            else
            {
                rowCount = dataGridViewNotMatchedInvoice.Rows.Count;
            }

            // Update label text
            invoiceAbsentInPDFCount.Text = $"Invoice Not Found In PDF Count: {rowCount}";
        }

        private void UpdateRowCountLabelForCompareExcel()
        {

            if (dataGridViewMatchedInvoice.DataSource == null)
            {
                compareExcelCount.Text = "No data loaded.";
                return;
            }
            
            int rowCount = 0;

            if (dataGridViewMatchedInvoice.DataSource is DataTable dataTable)
            {
                rowCount = dataTable.Rows.Count;
            }
            else if (dataGridViewMatchedInvoice.DataSource is BindingSource bindingSource)
            {
                rowCount = bindingSource.Count;
            }
            else if (dataGridViewMatchedInvoice.DataSource is IList list)
            {
                rowCount = list.Count;
            }
            else
            {
                rowCount = dataGridViewMatchedInvoice.Rows.Count;
            }

            // Update label text
            compareExcelCount.Text = $"Compare Excel Count: {rowCount}";
        }
        private void UpdateRowCountLabel()
        {

            if (dataGridViewForSourceExcel.DataSource == null)
            {
                sourceExcelCount.Text = "No data loaded.";
                return;
            } 

            int rowCount = 0;

            if (dataGridViewForSourceExcel.DataSource is DataTable dataTable)
            {
                rowCount = dataTable.Rows.Count;
            }
            else if (dataGridViewForSourceExcel.DataSource is BindingSource bindingSource)
            {
                rowCount = bindingSource.Count;
            }
            else if (dataGridViewForSourceExcel.DataSource is IList list)
            {
                rowCount = list.Count;
            }
            else
            {
                rowCount = dataGridViewForSourceExcel.Rows.Count;
            }

            // Update label text
            sourceExcelCount.Text = $"Source Excel Count: {rowCount}";
        }
        private async void LoadDataIntoFourthTab()
        {
            string inputText = selectedFolderPath;
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string pdfListPath = Path.Combine(desktopPath, pdfPath);

            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel");

            // Retrieve the SelectedIndex on the UI thread
            int selectedIndex = tabControl1.SelectedIndex;

            await Task.Run(() =>
            {
                // Ensure the destination directory exists
                if (Directory.Exists(destinationFolderPath))
                {
                    this.Invoke(new Action(() =>
                    {
                        dataGridViewNotMatchedInvoiceLoader.Visible = true;
                        dataGridViewNotMatchedInvoice.Visible = false;
                        invoiceAbsentInPDFCount.Text = "";
                    }));

                    // Get all PDF files in the directory 
                    string[] pdfFiles = Directory.GetFiles(inputText, "*.pdf");
                    if (pdfFiles.Length > 0)
                    {
                        string destinationFilePath = Path.Combine(destinationFolderPath, excelName);
                        PdfExtractionResult result = PdfExtractor.ExtractInvoicesFromPdfs(pdfFiles.ToList(), selectedIndex);

                        // Update the DataGridView on the UI thread
                        this.Invoke(new Action(() =>
                        {
                            dataGridViewNotMatchedInvoice.DataSource = result.InvoiceNotMatchedTable;
                            dataGridViewNotMatchedInvoiceLoader.Visible = false;
                            dataGridViewNotMatchedInvoice.Visible = true;

                            AddOrUpdateDataGridView("invoiceNotFoundInPDF", dataGridViewNotMatchedInvoice);
                            UpdateRowCountLabelForInvoiceAbsentPDF();
                        }));
                    }
                    else
                    {
                        // Show the message on the UI thread
                        this.Invoke(new Action(() => MessageBox.Show("Directory does not have any PDF files.")));
                    }
                }
                else
                {
                    // Show the message on the UI thread and create the directory
                    this.Invoke(new Action(() =>
                    {
                        MessageBox.Show("PDF Directory does not exist.");
                        Directory.CreateDirectory(destinationFolderPath);
                    }));
                }
            });
        }



        private async void LoadDataIntoFifthTab()
        {
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel");

            await Task.Run(() =>
            {
                // Ensure the destination directory exists
                if (Directory.Exists(destinationFolderPath))
                {
                    this.Invoke(new Action(() =>
                    {
                        dataGridICSAVaryingAddressLoader.Visible = true;
                        dataGridICSAVaryingAddress.Visible = false;
                    }));
                    PdfExtractionResult result = PdfExtractor.ExtractICSAAndAddress();

                    // Update the DataGridView on the UI thread
                    this.Invoke(new Action(() =>
                    {
                        dataGridICSAVaryingAddress.DataSource = result.ICSAVaryingAddress;
                        dataGridICSAVaryingAddressLoader.Visible = false;
                        dataGridICSAVaryingAddress.Visible = true;
                        
                        AddOrUpdateDataGridView("icsaExcel", dataGridICSAVaryingAddress);
                    }));
                }
                else
                {
                    // Show the message on the UI thread and create the directory
                    this.Invoke(new Action(() =>
                    {
                        MessageBox.Show("PDF Directory does not exist.");
                        Directory.CreateDirectory(destinationFolderPath);
                    }));
                }
            });
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {  
        }

        private async void LoadDataIntoTab()
        {
            string inputText = selectedFolderPath;
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string pdfListPath = Path.Combine(desktopPath, pdfPath);

            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel");
            string destinationFilePath = Path.Combine(destinationFolderPath, excelName);

            if (!File.Exists(destinationFilePath))
            {
                MessageBox.Show("Please upload the excel file again.");
                return; // Exit the method
            }

            if (!Directory.Exists(inputText))
            {
                MessageBox.Show("PDF Directory does not exist.");
                return; // Exit the method
            }

            // Run the file processing in a background task
            await Task.Run(() =>
            {
                string[] pdfFiles = Directory.GetFiles(inputText, "*.pdf");

                if (pdfFiles.Length > 0)
                {
                    PdfExtractionResult result = null;

                    // Retrieve the selected index on the UI thread
                    int selectedIndex = 0;
                    if (tabControl1.InvokeRequired)
                    {
                        tabControl1.Invoke(new Action(() => selectedIndex = tabControl1.SelectedIndex));
                    }
                    else
                    {
                        selectedIndex = tabControl1.SelectedIndex;
                    }

                    // Update UI controls based on the selected index
                    switch (selectedIndex)
                    {
                        case 1:
                            // Show the loader and hide the DataGridView on the UI thread
                            this.Invoke(new Action(() =>
                            {
                                dataGridViewMatchedInvoiceLoader.Visible = true;
                                dataGridViewMatchedInvoice.Visible = false;
                                compareExcelCount.Text = "";
                            }));

                            result = PdfExtractor.ExtractValuesFromPdfs(pdfFiles.ToList(), 1);

                            // Update the DataGridView on the UI thread
                            this.Invoke(new Action(() =>
                            {
                                dataGridViewMatchedInvoice.DataSource = result.MatchedTable;
                                dataGridViewMatchedInvoiceLoader.Visible = false;
                                dataGridViewMatchedInvoice.Visible = true;
                               
                                AddOrUpdateDataGridView("compareExcel", dataGridViewMatchedInvoice);
                                UpdateRowCountLabelForCompareExcel();
                            }));
                            
                            break;
                        case 2:
                            // Show the loader and hide the DataGridView on the UI thread
                            this.Invoke(new Action(() =>
                            {
                                invoiceButNotMatchedDtLoader.Visible = true;
                                invoiceButNotMatchedDt.Visible = false;
                                invoiceNotFoundExcelCount.Text = "";
                            }));

                            result = PdfExtractor.ExtractGrandTotalFromPdfs(pdfFiles.ToList(), 2);

                            // Update the DataGridView on the UI thread
                            this.Invoke(new Action(() =>
                            {
                                invoiceButNotMatchedDt.DataSource = result.InvoiceNotMatchedTable;
                                invoiceButNotMatchedDtLoader.Visible = false;
                                invoiceButNotMatchedDt.Visible = true;
                                UpdateRowCountLabelForInvoiceNotFoundInExcel();
                                AddOrUpdateDataGridView("invoiceNotFoundInExcel", invoiceButNotMatchedDt);
                                
                            }));
                            break;
                        case 4:
                            // Show the loader and hide the DataGridView on the UI thread
                            this.Invoke(new Action(() =>
                            {
                                listBoxLoader.Visible = true;
                                listBox.Visible = false;
                            }));

                            // listBoxLoader
                            result = PdfExtractor.ExtractGrandTotalFromPdfs(pdfFiles.ToList(), 4);

                            // Update the ListBox on the UI thread
                            this.Invoke(new Action(() =>
                            {
                                listBox.DataSource = result.ListOfPDFNotHavingInvoice;
                                listBoxLoader.Visible = false;
                                listBox.Visible = true;
                            }));
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    // Show the message on the UI thread
                    this.Invoke(new Action(() => MessageBox.Show("Directory does not have any PDF files.")));
                }
            });
        }




        // Event handler for TabControl's SelectedIndexChanged event
        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt=new DataTable();
            dataGridViewMatchedInvoice.DataSource = dt;
            dataGridICSAVaryingAddress.DataSource = dt;
            dataGridViewForSourceExcel.DataSource = dt; 
            dataGridViewNotMatchedInvoice.DataSource = dt;
            invoiceButNotMatchedDt.DataSource = dt;
            // Check if the second tab is selected

            switch (tabControl1.SelectedIndex)
            {
                case 0: 
                    LoadDataIntoFirstTab();
                    break; 
                case 1:
                case 2:
                case 4: 
                    if (Directory.Exists(selectedFolderPath))
                    {
                        LoadDataIntoTab();
                    }
                    else
                    {

                        DialogResult result = MessageBox.Show("Please select the PDF directory otherwise you will be redirected to first tab.", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result == DialogResult.OK)
                        {
                            tabControl1.SelectedIndex = 0;
                        }
                    }
                   
                    break;
                case 3:

                    if (Directory.Exists(selectedFolderPath))
                    {
                         LoadDataIntoFourthTab();
                    }
                    else
                    {

                        DialogResult result = MessageBox.Show("Please select the PDF directory.", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result == DialogResult.OK)
                        {
                            tabControl1.SelectedIndex = 0;
                        }
                    }

                    break;
                case 5:
                    LoadDataIntoFifthTab();
                    break;
                default: 
                    break;
            }

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls|*.pdf|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string sourceFilePath = openFileDialog.FileName;
                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
                string destinationFolderPath = Path.Combine(basePath, "UploadedExcel");

                // Ensure the destination directory exists
                if (!Directory.Exists(destinationFolderPath))
                {
                    Directory.CreateDirectory(destinationFolderPath);
                }

                string destinationFilePath = Path.Combine(destinationFolderPath, excelName);

                try
                {
                    // Check if source file exists
                    if (!File.Exists(sourceFilePath))
                    {
                        MessageBox.Show("Source file does not exist: " + sourceFilePath);
                        return;
                    }

                    // If the file already exists in the destination, delete it
                    if (File.Exists(destinationFilePath))
                    {
                        File.Delete(destinationFilePath);
                        dataGridViewForSourceExcel.DataSource = null;
                    }

                    // Copy the file to the destination directory
                    File.Copy(sourceFilePath, destinationFilePath, true);
                    tabControl1.Visible = true;
                    dataGridViewForSourceExcel.DataSource = new DataTable();
                    LoadDataIntoFirstTab();
                    this.uploadButton.Text = "Replace Excel";

                    // Ensure the file was copied successfully
                    if (!File.Exists(destinationFilePath))
                    {
                        MessageBox.Show("File copy failed.");
                        return;
                    }
                     

                    MessageBox.Show("File uploaded and copied successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}");
                }
            }
        }

        private void AddOrUpdateDataGridView(string key, DataGridView dataGridView)
        {
            // Check if the key already exists in the dictionary
            if (dataGridViewDictionary.ContainsKey(key))
            {
                // Remove the existing entry with the same key
                dataGridViewDictionary.Remove(key);
            }

            // Add the new DataGridView to the dictionary
            dataGridViewDictionary.Add(key, dataGridView);
        }
        private void exportsExcel_Click_1(object sender, EventArgs e)
        {
            DataGridView sourceExcelDGV;
            DataGridView icsaExcelDGV;
            DataGridView compareExcelDGV;
            DataGridView invoiceNotFoundInExcelDGV;
            DataGridView invoiceNotFoundInPDFDGV;

            // Retrieve DataGridViews from the dictionary
            dataGridViewDictionary.TryGetValue("sourceExcel", out sourceExcelDGV);
            dataGridViewDictionary.TryGetValue("icsaExcel", out icsaExcelDGV);
            dataGridViewDictionary.TryGetValue("compareExcel", out compareExcelDGV);
            dataGridViewDictionary.TryGetValue("invoiceNotFoundInExcel", out invoiceNotFoundInExcelDGV);
            dataGridViewDictionary.TryGetValue("invoiceNotFoundInPDF", out invoiceNotFoundInPDFDGV);

            DataGridView[] dataGridViews = { sourceExcelDGV,      compareExcelDGV,
                invoiceNotFoundInExcelDGV,
                invoiceNotFoundInPDFDGV,
                icsaExcelDGV }; // Add your DataGridViews here
            string[] sheetNames = { "Source Excel", "Compare PDF excel", "Invoice not found in Excel", "Invoice Not found in PDF", "ICSA with Varying Address" }; // Corresponding sheet names

            // Get the path to the Downloads folder
            string downloadsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string filePath = Path.Combine(downloadsFolderPath, "ExportedData.xlsx");

            ExportDataGridViewsToExcel(dataGridViews, sheetNames, filePath);
          

        }

        

        public static void ExportDataGridViewsToExcel(DataGridView[] dataGridViews, string[] sheetNames, string filePath)
        {
            if (dataGridViews.Length != sheetNames.Length)
            {
                throw new ArgumentException("The number of DataGridViews must match the number of sheet names.");
            }
            // Load existing workbook or create a new one
            XLWorkbook workbook;

            try
            {
                // Check if the file exists and load it or create a new one
                if (System.IO.File.Exists(filePath))
                {
                    workbook = new XLWorkbook(filePath);
                }
                else
                {
                    workbook = new XLWorkbook();
                }
            }
            catch (System.IO.IOException ex)
            {
                // Show a message box to the user indicating the error
                MessageBox.Show("The process cannot access the file because it is being used by another process. Please close any applications that may be using the file and try again.", "File Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // Exit the method if the file cannot be accessed
            }

             
            // Add new sheets from DataGridViews
            for (int i = 0; i < dataGridViews.Length; i++)
            {
                var dataGridView = dataGridViews[i];
                var sheetName = sheetNames[i];
                if (dataGridView != null)
                {
                    DataTable dt = GetDataTableFromDataGridView(dataGridView);
                    if (dt.Rows.Count > 0)
                    {
                        if (workbook.Worksheets.Count >0 && workbook.Worksheets.Contains(sheetName))
                        {
                            workbook.Worksheets.Delete(sheetName); 
                        }
                        var worksheet = workbook.Worksheets.Add(sheetName);
                        worksheet.Cell(1, 1).InsertTable(dt);
                    }
                }
            }

            // Save the workbook
            workbook.SaveAs(filePath);
            MessageBox.Show($"Excel is downloaded in folder {filePath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }

        private static DataTable GetDataTableFromDataGridView(DataGridView dgv)
        {
            var dt = new DataTable();

            // Check if the DataGridView has a data source
            if (dgv.DataSource is DataTable dataSourceTable)
            {
                // If the DataGridView is bound to a DataTable, return it directly
                return dataSourceTable.Copy();
            }
            else if (dgv.DataSource != null)
            {
                // Handle other data sources (e.g., List<T>, BindingList<T>)
                // You could add logic here to convert them to DataTable if needed
                throw new InvalidOperationException("Unsupported data source type.");
            }

            // Add columns to the DataTable
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType ?? typeof(string));
            }

            // Add rows to the DataTable
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue; // Skip the new row placeholder

                var dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value ?? DBNull.Value; // Use DBNull.Value for null cells
                }
                dt.Rows.Add(dataRow);
            }

            return dt;
        }
         
        private void button2_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a folder";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected folder path
                    selectedFolderPath = folderBrowserDialog.SelectedPath;

                    // Display the selected folder path (or use it as needed)
                    MessageBox.Show("Selected folder: " + selectedFolderPath, "Folder Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
         
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridViewMatchedInvoice_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage1_Click_1(object sender, EventArgs e)
        {

        }

        private void invoiceButNotMatchedDt_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
 

        private void tabPage1_Click_2(object sender, EventArgs e)
        {

        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click_3(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

       
    }
}
