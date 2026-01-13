using System.Configuration;
using System;
using System.IO;
using System.Windows.Forms;

namespace PdfReader
{
    partial class PdfScan
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
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.uploadButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.folderBrowserDialog = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.sourceExcelTab = new System.Windows.Forms.TabPage();
            this.loaderLabel = new System.Windows.Forms.Label();
            this.sourceExcelCount = new System.Windows.Forms.Label();
            this.dataGridViewForSourceExcel = new System.Windows.Forms.DataGridView();
            this.compareExcelTab = new System.Windows.Forms.TabPage();
            this.compareExcelCount = new System.Windows.Forms.Label();
            this.dataGridViewMatchedInvoiceLoader = new System.Windows.Forms.Label();
            this.dataGridViewMatchedInvoice = new System.Windows.Forms.DataGridView();
            this.invoiceNotMatchedTab = new System.Windows.Forms.TabPage();
            this.invoiceNotFoundExcelCount = new System.Windows.Forms.Label();
            this.invoiceButNotMatchedDtLoader = new System.Windows.Forms.Label();
            this.invoiceButNotMatchedDt = new System.Windows.Forms.DataGridView();
            this.excelNotHavingInvoiceFromPDF = new System.Windows.Forms.TabPage();
            this.invoiceAbsentInPDFCount = new System.Windows.Forms.Label();
            this.dataGridViewNotMatchedInvoiceLoader = new System.Windows.Forms.Label();
            this.dataGridViewNotMatchedInvoice = new System.Windows.Forms.DataGridView();
            this.pdfNotHavingInvoice = new System.Windows.Forms.TabPage();
            this.listBoxLoader = new System.Windows.Forms.Label();
            this.listBox = new System.Windows.Forms.ListBox();
            this.icsaWithVaringAddress = new System.Windows.Forms.TabPage();
            this.dataGridICSAVaryingAddressLoader = new System.Windows.Forms.Label();
            this.dataGridICSAVaryingAddress = new System.Windows.Forms.DataGridView();
            this.exportsExcel = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.sourceExcelTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewForSourceExcel)).BeginInit();
            this.compareExcelTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMatchedInvoice)).BeginInit();
            this.invoiceNotMatchedTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.invoiceButNotMatchedDt)).BeginInit();
            this.excelNotHavingInvoiceFromPDF.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNotMatchedInvoice)).BeginInit();
            this.pdfNotHavingInvoice.SuspendLayout();
            this.icsaWithVaringAddress.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridICSAVaryingAddress)).BeginInit();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(314, 158);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(6, 6);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // uploadButton
            // 
            this.uploadButton.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.uploadButton.Location = new System.Drawing.Point(757, 32);
            this.uploadButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.uploadButton.Name = "uploadButton";
            this.uploadButton.Size = new System.Drawing.Size(418, 43);
            this.uploadButton.TabIndex = 0;
            this.uploadButton.Text = "Upload Excel";
            this.uploadButton.UseVisualStyleBackColor = false;
            this.uploadButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 78.8764F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.1236F));
            this.tableLayoutPanel1.Location = new System.Drawing.Point(45, -1);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(4);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 40.96386F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 59.03614F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 341F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(0, 341);
            this.tableLayoutPanel1.TabIndex = 5;
            // 
            // folderBrowserDialog
            // 
            this.folderBrowserDialog.ForeColor = System.Drawing.SystemColors.Highlight;
            this.folderBrowserDialog.Location = new System.Drawing.Point(23, 32);
            this.folderBrowserDialog.Name = "folderBrowserDialog";
            this.folderBrowserDialog.Size = new System.Drawing.Size(381, 43);
            this.folderBrowserDialog.TabIndex = 18;
            this.folderBrowserDialog.Text = "Choose PDF Folder";
            this.folderBrowserDialog.UseVisualStyleBackColor = true;
            this.folderBrowserDialog.Click += new System.EventHandler(this.button2_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.sourceExcelTab);
            this.tabControl1.Controls.Add(this.compareExcelTab);
            this.tabControl1.Controls.Add(this.invoiceNotMatchedTab);
            this.tabControl1.Controls.Add(this.excelNotHavingInvoiceFromPDF);
            this.tabControl1.Controls.Add(this.pdfNotHavingInvoice);
            this.tabControl1.Controls.Add(this.icsaWithVaringAddress);
            this.tabControl1.Cursor = System.Windows.Forms.Cursors.Default;
            this.tabControl1.Location = new System.Drawing.Point(23, 93);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1224, 552);
            this.tabControl1.TabIndex = 19;
            // 
            // sourceExcelTab
            // 
            this.sourceExcelTab.Controls.Add(this.loaderLabel);
            this.sourceExcelTab.Controls.Add(this.sourceExcelCount);
            this.sourceExcelTab.Controls.Add(this.dataGridViewForSourceExcel);
            this.sourceExcelTab.Location = new System.Drawing.Point(4, 25);
            this.sourceExcelTab.Name = "sourceExcelTab";
            this.sourceExcelTab.Padding = new System.Windows.Forms.Padding(3);
            this.sourceExcelTab.Size = new System.Drawing.Size(1216, 523);
            this.sourceExcelTab.TabIndex = 0;
            this.sourceExcelTab.Text = "Source Excel";
            this.sourceExcelTab.UseVisualStyleBackColor = true;
            this.sourceExcelTab.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // loaderLabel
            // 
            this.loaderLabel.AutoSize = true;
            this.loaderLabel.Location = new System.Drawing.Point(411, 11);
            this.loaderLabel.Name = "loaderLabel";
            this.loaderLabel.Size = new System.Drawing.Size(65, 16);
            this.loaderLabel.TabIndex = 2;
            this.loaderLabel.Text = "Loading...";
            this.loaderLabel.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // sourceExcelCount
            // 
            this.sourceExcelCount.AutoSize = true;
            this.sourceExcelCount.ForeColor = System.Drawing.SystemColors.ControlText;
            this.sourceExcelCount.Location = new System.Drawing.Point(8, 11);
            this.sourceExcelCount.Name = "sourceExcelCount";
            this.sourceExcelCount.Size = new System.Drawing.Size(123, 16);
            this.sourceExcelCount.TabIndex = 1;
            this.sourceExcelCount.Text = "Source Excel Count";
            this.sourceExcelCount.Click += new System.EventHandler(this.label1_Click);
            // 
            // dataGridViewForSourceExcel
            // 
            this.dataGridViewForSourceExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewForSourceExcel.Location = new System.Drawing.Point(8, 40);
            this.dataGridViewForSourceExcel.Name = "dataGridViewForSourceExcel";
            this.dataGridViewForSourceExcel.RowHeadersWidth = 51;
            this.dataGridViewForSourceExcel.RowTemplate.Height = 24;
            this.dataGridViewForSourceExcel.Size = new System.Drawing.Size(1186, 455);
            this.dataGridViewForSourceExcel.TabIndex = 0;
            this.dataGridViewForSourceExcel.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // compareExcelTab
            // 
            this.compareExcelTab.Controls.Add(this.compareExcelCount);
            this.compareExcelTab.Controls.Add(this.dataGridViewMatchedInvoiceLoader);
            this.compareExcelTab.Controls.Add(this.dataGridViewMatchedInvoice);
            this.compareExcelTab.Cursor = System.Windows.Forms.Cursors.Default;
            this.compareExcelTab.Location = new System.Drawing.Point(4, 25);
            this.compareExcelTab.Name = "compareExcelTab";
            this.compareExcelTab.Padding = new System.Windows.Forms.Padding(3);
            this.compareExcelTab.Size = new System.Drawing.Size(1216, 523);
            this.compareExcelTab.TabIndex = 1;
            this.compareExcelTab.Text = "Compare PDF with Excel";
            this.compareExcelTab.UseVisualStyleBackColor = true;
            this.compareExcelTab.Click += new System.EventHandler(this.tabPage2_Click);
            // 
            // compareExcelCount
            // 
            this.compareExcelCount.AutoSize = true;
            this.compareExcelCount.ForeColor = System.Drawing.SystemColors.ControlText;
            this.compareExcelCount.Location = new System.Drawing.Point(14, 13);
            this.compareExcelCount.Name = "compareExcelCount";
            this.compareExcelCount.Size = new System.Drawing.Size(136, 16);
            this.compareExcelCount.TabIndex = 7;
            this.compareExcelCount.Text = "Compare Excel Count";
            // 
            // dataGridViewMatchedInvoiceLoader
            // 
            this.dataGridViewMatchedInvoiceLoader.AutoSize = true;
            this.dataGridViewMatchedInvoiceLoader.Location = new System.Drawing.Point(412, 13);
            this.dataGridViewMatchedInvoiceLoader.Name = "dataGridViewMatchedInvoiceLoader";
            this.dataGridViewMatchedInvoiceLoader.Size = new System.Drawing.Size(65, 16);
            this.dataGridViewMatchedInvoiceLoader.TabIndex = 6;
            this.dataGridViewMatchedInvoiceLoader.Text = "Loading...";
            // 
            // dataGridViewMatchedInvoice
            // 
            this.dataGridViewMatchedInvoice.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridViewMatchedInvoice.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridViewMatchedInvoice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMatchedInvoice.Location = new System.Drawing.Point(12, 40);
            this.dataGridViewMatchedInvoice.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridViewMatchedInvoice.Name = "dataGridViewMatchedInvoice";
            this.dataGridViewMatchedInvoice.RowHeadersWidth = 51;
            this.dataGridViewMatchedInvoice.Size = new System.Drawing.Size(1147, 458);
            this.dataGridViewMatchedInvoice.TabIndex = 5;
            this.dataGridViewMatchedInvoice.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMatchedInvoice_CellContentClick);
            // 
            // invoiceNotMatchedTab
            // 
            this.invoiceNotMatchedTab.Controls.Add(this.invoiceNotFoundExcelCount);
            this.invoiceNotMatchedTab.Controls.Add(this.invoiceButNotMatchedDtLoader);
            this.invoiceNotMatchedTab.Controls.Add(this.invoiceButNotMatchedDt);
            this.invoiceNotMatchedTab.Location = new System.Drawing.Point(4, 25);
            this.invoiceNotMatchedTab.Name = "invoiceNotMatchedTab";
            this.invoiceNotMatchedTab.Padding = new System.Windows.Forms.Padding(3);
            this.invoiceNotMatchedTab.Size = new System.Drawing.Size(1216, 523);
            this.invoiceNotMatchedTab.TabIndex = 2;
            this.invoiceNotMatchedTab.Text = "Invoice Not Found In Excel";
            this.invoiceNotMatchedTab.UseVisualStyleBackColor = true;
            this.invoiceNotMatchedTab.Click += new System.EventHandler(this.tabPage1_Click_1);
            // 
            // invoiceNotFoundExcelCount
            // 
            this.invoiceNotFoundExcelCount.AutoSize = true;
            this.invoiceNotFoundExcelCount.Location = new System.Drawing.Point(18, 12);
            this.invoiceNotFoundExcelCount.Name = "invoiceNotFoundExcelCount";
            this.invoiceNotFoundExcelCount.Size = new System.Drawing.Size(188, 16);
            this.invoiceNotFoundExcelCount.TabIndex = 11;
            this.invoiceNotFoundExcelCount.Text = "Invoice Not Found Excel Count";
            // 
            // invoiceButNotMatchedDtLoader
            // 
            this.invoiceButNotMatchedDtLoader.AutoSize = true;
            this.invoiceButNotMatchedDtLoader.Location = new System.Drawing.Point(427, 10);
            this.invoiceButNotMatchedDtLoader.Name = "invoiceButNotMatchedDtLoader";
            this.invoiceButNotMatchedDtLoader.Size = new System.Drawing.Size(65, 16);
            this.invoiceButNotMatchedDtLoader.TabIndex = 10;
            this.invoiceButNotMatchedDtLoader.Text = "Loading...";
            // 
            // invoiceButNotMatchedDt
            // 
            this.invoiceButNotMatchedDt.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.invoiceButNotMatchedDt.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.invoiceButNotMatchedDt.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.invoiceButNotMatchedDt.Location = new System.Drawing.Point(16, 40);
            this.invoiceButNotMatchedDt.Name = "invoiceButNotMatchedDt";
            this.invoiceButNotMatchedDt.RowHeadersWidth = 51;
            this.invoiceButNotMatchedDt.RowTemplate.Height = 24;
            this.invoiceButNotMatchedDt.Size = new System.Drawing.Size(1137, 461);
            this.invoiceButNotMatchedDt.TabIndex = 9;
            this.invoiceButNotMatchedDt.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.invoiceButNotMatchedDt_CellContentClick);
            // 
            // excelNotHavingInvoiceFromPDF
            // 
            this.excelNotHavingInvoiceFromPDF.Controls.Add(this.invoiceAbsentInPDFCount);
            this.excelNotHavingInvoiceFromPDF.Controls.Add(this.dataGridViewNotMatchedInvoiceLoader);
            this.excelNotHavingInvoiceFromPDF.Controls.Add(this.dataGridViewNotMatchedInvoice);
            this.excelNotHavingInvoiceFromPDF.Location = new System.Drawing.Point(4, 25);
            this.excelNotHavingInvoiceFromPDF.Name = "excelNotHavingInvoiceFromPDF";
            this.excelNotHavingInvoiceFromPDF.Padding = new System.Windows.Forms.Padding(3);
            this.excelNotHavingInvoiceFromPDF.Size = new System.Drawing.Size(1216, 523);
            this.excelNotHavingInvoiceFromPDF.TabIndex = 3;
            this.excelNotHavingInvoiceFromPDF.Text = "Invoice Not Found in PDF";
            this.excelNotHavingInvoiceFromPDF.UseVisualStyleBackColor = true;
            // 
            // invoiceAbsentInPDFCount
            // 
            this.invoiceAbsentInPDFCount.AutoSize = true;
            this.invoiceAbsentInPDFCount.Location = new System.Drawing.Point(13, 14);
            this.invoiceAbsentInPDFCount.Name = "invoiceAbsentInPDFCount";
            this.invoiceAbsentInPDFCount.Size = new System.Drawing.Size(201, 16);
            this.invoiceAbsentInPDFCount.TabIndex = 10;
            this.invoiceAbsentInPDFCount.Text = "Invoice Not Found In Excel Count";
            // 
            // dataGridViewNotMatchedInvoiceLoader
            // 
            this.dataGridViewNotMatchedInvoiceLoader.AutoSize = true;
            this.dataGridViewNotMatchedInvoiceLoader.Location = new System.Drawing.Point(433, 12);
            this.dataGridViewNotMatchedInvoiceLoader.Name = "dataGridViewNotMatchedInvoiceLoader";
            this.dataGridViewNotMatchedInvoiceLoader.Size = new System.Drawing.Size(65, 16);
            this.dataGridViewNotMatchedInvoiceLoader.TabIndex = 9;
            this.dataGridViewNotMatchedInvoiceLoader.Text = "Loading...";
            // 
            // dataGridViewNotMatchedInvoice
            // 
            this.dataGridViewNotMatchedInvoice.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridViewNotMatchedInvoice.ColumnHeadersHeight = 29;
            this.dataGridViewNotMatchedInvoice.Location = new System.Drawing.Point(7, 40);
            this.dataGridViewNotMatchedInvoice.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridViewNotMatchedInvoice.Name = "dataGridViewNotMatchedInvoice";
            this.dataGridViewNotMatchedInvoice.RowHeadersWidth = 51;
            this.dataGridViewNotMatchedInvoice.Size = new System.Drawing.Size(1172, 461);
            this.dataGridViewNotMatchedInvoice.TabIndex = 8;
            // 
            // pdfNotHavingInvoice
            // 
            this.pdfNotHavingInvoice.Controls.Add(this.listBoxLoader);
            this.pdfNotHavingInvoice.Controls.Add(this.listBox);
            this.pdfNotHavingInvoice.Location = new System.Drawing.Point(4, 25);
            this.pdfNotHavingInvoice.Name = "pdfNotHavingInvoice";
            this.pdfNotHavingInvoice.Padding = new System.Windows.Forms.Padding(3);
            this.pdfNotHavingInvoice.Size = new System.Drawing.Size(1216, 523);
            this.pdfNotHavingInvoice.TabIndex = 4;
            this.pdfNotHavingInvoice.Text = "Invalid PDF";
            this.pdfNotHavingInvoice.UseVisualStyleBackColor = true;
            this.pdfNotHavingInvoice.Click += new System.EventHandler(this.tabPage1_Click_2);
            // 
            // listBoxLoader
            // 
            this.listBoxLoader.AutoSize = true;
            this.listBoxLoader.Location = new System.Drawing.Point(406, 20);
            this.listBoxLoader.Name = "listBoxLoader";
            this.listBoxLoader.Size = new System.Drawing.Size(65, 16);
            this.listBoxLoader.TabIndex = 9;
            this.listBoxLoader.Text = "Loading...";
            // 
            // listBox
            // 
            this.listBox.FormattingEnabled = true;
            this.listBox.ItemHeight = 16;
            this.listBox.Location = new System.Drawing.Point(30, 49);
            this.listBox.Name = "listBox";
            this.listBox.Size = new System.Drawing.Size(879, 260);
            this.listBox.TabIndex = 8;
            this.listBox.SelectedIndexChanged += new System.EventHandler(this.listBox_SelectedIndexChanged);
            // 
            // icsaWithVaringAddress
            // 
            this.icsaWithVaringAddress.Controls.Add(this.dataGridICSAVaryingAddressLoader);
            this.icsaWithVaringAddress.Controls.Add(this.dataGridICSAVaryingAddress);
            this.icsaWithVaringAddress.Location = new System.Drawing.Point(4, 25);
            this.icsaWithVaringAddress.Name = "icsaWithVaringAddress";
            this.icsaWithVaringAddress.Padding = new System.Windows.Forms.Padding(3);
            this.icsaWithVaringAddress.Size = new System.Drawing.Size(1216, 523);
            this.icsaWithVaringAddress.TabIndex = 5;
            this.icsaWithVaringAddress.Text = "ICSA with varying addresses";
            this.icsaWithVaringAddress.UseVisualStyleBackColor = true;
            this.icsaWithVaringAddress.Click += new System.EventHandler(this.tabPage1_Click_3);
            // 
            // dataGridICSAVaryingAddressLoader
            // 
            this.dataGridICSAVaryingAddressLoader.AutoSize = true;
            this.dataGridICSAVaryingAddressLoader.Location = new System.Drawing.Point(420, 11);
            this.dataGridICSAVaryingAddressLoader.Name = "dataGridICSAVaryingAddressLoader";
            this.dataGridICSAVaryingAddressLoader.Size = new System.Drawing.Size(65, 16);
            this.dataGridICSAVaryingAddressLoader.TabIndex = 1;
            this.dataGridICSAVaryingAddressLoader.Text = "Loading...";
            // 
            // dataGridICSAVaryingAddress
            // 
            this.dataGridICSAVaryingAddress.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridICSAVaryingAddress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridICSAVaryingAddress.Location = new System.Drawing.Point(12, 40);
            this.dataGridICSAVaryingAddress.Name = "dataGridICSAVaryingAddress";
            this.dataGridICSAVaryingAddress.RowHeadersWidth = 51;
            this.dataGridICSAVaryingAddress.RowTemplate.Height = 24;
            this.dataGridICSAVaryingAddress.Size = new System.Drawing.Size(1196, 452);
            this.dataGridICSAVaryingAddress.TabIndex = 0;
            // 
            // exportsExcel
            // 
            this.exportsExcel.Location = new System.Drawing.Point(432, 32);
            this.exportsExcel.Name = "exportsExcel";
            this.exportsExcel.Size = new System.Drawing.Size(309, 42);
            this.exportsExcel.TabIndex = 20;
            this.exportsExcel.Text = "Exports Excel";
            this.exportsExcel.UseVisualStyleBackColor = true;
            this.exportsExcel.Click += new System.EventHandler(this.exportsExcel_Click_1);
            // 
            // PdfScan
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1239, 1076);
            this.Controls.Add(this.exportsExcel);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.folderBrowserDialog);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.uploadButton);
            this.Controls.Add(this.richTextBox1);
            this.ForeColor = System.Drawing.SystemColors.Highlight;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "PdfScan";
            this.Text = "Pdf Scan";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.sourceExcelTab.ResumeLayout(false);
            this.sourceExcelTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewForSourceExcel)).EndInit();
            this.compareExcelTab.ResumeLayout(false);
            this.compareExcelTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMatchedInvoice)).EndInit();
            this.invoiceNotMatchedTab.ResumeLayout(false);
            this.invoiceNotMatchedTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.invoiceButNotMatchedDt)).EndInit();
            this.excelNotHavingInvoiceFromPDF.ResumeLayout(false);
            this.excelNotHavingInvoiceFromPDF.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNotMatchedInvoice)).EndInit();
            this.pdfNotHavingInvoice.ResumeLayout(false);
            this.pdfNotHavingInvoice.PerformLayout();
            this.icsaWithVaringAddress.ResumeLayout(false);
            this.icsaWithVaringAddress.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridICSAVaryingAddress)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button uploadButton;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button folderBrowserDialog;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage sourceExcelTab;
        private System.Windows.Forms.TabPage compareExcelTab;
        private System.Windows.Forms.DataGridView dataGridViewForSourceExcel;
        private System.Windows.Forms.DataGridView dataGridViewMatchedInvoice;
        private TabPage invoiceNotMatchedTab;
        private DataGridView invoiceButNotMatchedDt;
        private TabPage excelNotHavingInvoiceFromPDF;
        private TabPage pdfNotHavingInvoice;
        private ListBox listBox;
        private DataGridView dataGridViewNotMatchedInvoice;
        private TabPage icsaWithVaringAddress;
        private DataGridView dataGridICSAVaryingAddress;
        private Label sourceExcelCount;
        private Label loaderLabel;
        private Label dataGridViewMatchedInvoiceLoader;
        private Label invoiceButNotMatchedDtLoader;
        private Label listBoxLoader;
        private Label dataGridViewNotMatchedInvoiceLoader;
        private Label dataGridICSAVaryingAddressLoader;
        private Label compareExcelCount;
        private Label invoiceNotFoundExcelCount;
        private Label invoiceAbsentInPDFCount;
        private Button exportsExcel;
    }
}

