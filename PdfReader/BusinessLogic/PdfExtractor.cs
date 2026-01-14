using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout.Element;
using iText.Layout.Properties.Grid;
using OfficeOpenXml;
using PdfReader.BusinessLogic.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks; 
using System.Windows.Forms; 

namespace PdfReader.BusinessLogic
{
    public class PdfExtractor
    { 
        public PdfExtractor() {
            // Set the EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        } 
        static List<string> headers = new List<string>
        {
            "Entity Name",
            "CSC NAME",
            "BUSINESS UNIT",
            "INVOICE NUMBER",
            "Sum of TOTAL INV AMOUNT INR",
            "Sum of TOTAL INV AMOUNT USD",
            "Grand Value",
            "Matched"
        };

        static List<string> headersForMatched = new List<string>
        {
            "Entity Name",
            "Bill to Details PDF", 
            "CSC NAME",
            "BUSINESS UNIT",
            "Services Month",
            "Invoice Date",
            "Invoice Date PDF", 
            "INVOICE NUMBER",
            "DOM / EXP",
            "LOCATION",
            "FX RATE",
            "FX RATE PDF",
            "Country Serviced",
            "Sum of TOTAL INV AMOUNT INR",
            "Grand Value PDF", 
            "Sum of TOTAL INV AMOUNT USD",
            "Address",
            "HSN",
            "ICSA",
            "Matched",
            "Notes"
        };


        static List<string> headerOfNotFoundInvoiceInExcel = new List<string>
        {
            "PDF",
            "INVOICE NUMBER",
            "Grand Value"
        };

        static List<string> headerOfICSAVaryingAddress = new List<string>
        {
            "INVOICE NUMBER",
            "ICSA",
            "Address"
        };
        public static string GetUploadedExcelPath()
        {
            string excelName = ConfigurationManager.AppSettings["ExcelName"];
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel", excelName);

            return destinationFolderPath;
        }
        public static PdfExtractionResult ExtractExcelTuple()
        { 
            DataTable matchedDt = new DataTable(); 
            string destinationFolderPath = GetUploadedExcelPath();
            ReadExcelFileAndFetchData(destinationFolderPath, matchedDt);
            return new PdfExtractionResult
            {
                SourceExcelTable = matchedDt
            };
        }
        
        public static PdfExtractionResult ExtractInvoicesFromPdfs(List<string> pdfPaths, int tabNumber)
        {
            var listOfInvoice = new List<string>(); 
            Dictionary<string, string> invoiceGrandValue = new Dictionary<string, string>();
            Dictionary<string, string> invoicePDFName = new Dictionary<string, string>();
            foreach (var pdfPath in pdfPaths)
            {
                if (File.Exists(pdfPath))
                {
                    var invoiceNumber = ExtractGrandTotalFromPdf(pdfPath, "Invoice No");
                    var grandValue = ExtractGrandTotalFromPdf(pdfPath, "Grand Total");
                    if (!string.IsNullOrEmpty(invoiceNumber))
                    {
                        listOfInvoice.Add(invoiceNumber); 
                    } 
                }

            } 
            DataTable invoiceButNotMatchedDt = new DataTable();
            string excelName = ConfigurationManager.AppSettings["ExcelName"];
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel", excelName);

            NotMatchedTupleIncExcelFile(destinationFolderPath, invoiceButNotMatchedDt, listOfInvoice);
            return new PdfExtractionResult
            { 
                InvoiceNotMatchedTable = invoiceButNotMatchedDt,
            };
        }

   
        public static string ExtractMonthFromPdf(string pdfPath)
        {
            // Update the problematic line:
             var pdfDocument = new PdfDocument(new iText.Kernel.Pdf.PdfReader(pdfPath));

            for (int pageNum = 1; pageNum <= pdfDocument.GetNumberOfPages(); pageNum++)
            {
                var page = pdfDocument.GetPage(pageNum);
                var strategy = new SimpleTextExtractionStrategy();
                var pageText = PdfTextExtractor.GetTextFromPage(page, strategy);

                var match = Regex.Match(
                    pageText,
                    @"month\s+of\s+([A-Za-z]{3}-\d{2})",
                    RegexOptions.IgnoreCase
                );

                if (match.Success)
                {
                    return match.Groups[1].Value; // Nov-25
                }
            }

            return string.Empty;
        }

        public static PdfExtractionResult ExtractValuesFromPdfs(List<string> pdfPaths, int tabNumber)
        {
            try
            {
                var listOfInvoice = new List<string>();
                Dictionary<string, PdfExtractedValue> invoiceGrandValue = new Dictionary<string, PdfExtractedValue>();
                foreach (var pdfPath in pdfPaths)
                {
                    if (File.Exists(pdfPath))
                    {
                        var invoiceType = ExtractTopTextFromPdf(pdfPath);
                        var invoiceNumber = ExtractGrandTotalFromPdf(pdfPath, "Invoice No. :");
                        var grandValue = ExtractGrandTotalFromPdf(pdfPath, "Grand Total");
                        var invoiceDateValue = ExtractGrandTotalFromPdf(pdfPath, "Invoice Date :");
                        var billToDetailsValue = ExtractGrandTotalFromPdf(pdfPath, "Bill To");
                        var exchangeRateValue = ExtractGrandTotalFromPdf(pdfPath, "Authorised Signatory");
                        var serviceDateValue = ExtractMonthFromPdf(pdfPath);
                        if (!string.IsNullOrEmpty(invoiceNumber))
                        {
                            PdfExtractedValue pdfExtractedValue = new PdfExtractedValue();
                            pdfExtractedValue.GrandTotalValue = grandValue;
                            pdfExtractedValue.InvoiceValue = invoiceNumber;
                            pdfExtractedValue.InvoiceDateValue = invoiceDateValue;
                            pdfExtractedValue.BillToDetailsValue = billToDetailsValue;
                            pdfExtractedValue.ExchangeRateValue = exchangeRateValue;
                            pdfExtractedValue.ServiceDateValue = serviceDateValue;
                            listOfInvoice.Add(invoiceNumber);
                            if (!invoiceGrandValue.ContainsKey(invoiceNumber))
                            {
                                invoiceGrandValue.Add(invoiceNumber, pdfExtractedValue);
                            }
                            else
                            {
                                // Update existing entry (if needed)
                                invoiceGrandValue[invoiceNumber] = pdfExtractedValue;
                            }
                        }
                    }
                }
                DataTable matchedDt = new DataTable();
                string destinationFolderPath = GetUploadedExcelPath();

                var duplicateHeaders = headersForMatched
                    .GroupBy(h => h)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

                if (duplicateHeaders.Any())
                {
                    MessageBox.Show("Duplicate headers: " + string.Join(", ", duplicateHeaders));
                }

                List<string> listOfPDFHavingInvoiceAndPrsentInExcel = new List<string>();

                foreach (var kvp in invoiceGrandValue)
                {
                    var invoice = kvp.Key;
                    var pdfData = kvp.Value;

                    CompareInvoiceInExcelFile(
                        destinationFolderPath,
                        matchedDt,
                        pdfData,
                        listOfPDFHavingInvoiceAndPrsentInExcel
                    );
                }

                return new PdfExtractionResult
                {
                    MatchedTable = matchedDt
                };
            }
            catch (Exception ex)
            { 
                MessageBox.Show(ex.ToString()); // shows exact line & key
                throw; // rethrow so VS breaks on correct line
                // Return an empty result in case of an exception
               
            }
        }

        public static PdfExtractionResult ExtractGrandTotalFromPdfs(List<string> pdfPaths, int tabNumber)
        {
            var listOfInvoice = new List<string>();
            var listOfPDFNotHavingInvoice = new List<string>();
            Dictionary<string, string> invoiceGrandValue = new Dictionary<string, string>();
            Dictionary<string, string> invoicePDFName = new Dictionary<string, string>();
            foreach (var pdfPath in pdfPaths)
            {
                if (File.Exists(pdfPath))
                {
                    var invoiceNumber = ExtractGrandTotalFromPdf(pdfPath, "Invoice No");
                    var grandValue = ExtractGrandTotalFromPdf(pdfPath, "Grand Total");
                    if (!string.IsNullOrEmpty(invoiceNumber))
                    { 
                      listOfInvoice.Add(invoiceNumber);
                      invoiceGrandValue.Add(invoiceNumber, grandValue);
                      invoicePDFName.Add(invoiceNumber, Path.GetFileName(pdfPath));
                    }
                    else
                    {
                        listOfPDFNotHavingInvoice.Add(pdfPath);
                    } 
                }
                 
            }
            DataTable matchedDt = new DataTable();
            DataTable notMatchedDt = new DataTable();
            DataTable invoiceButNotMatchedDt = new DataTable();
            string excelName = ConfigurationManager.AppSettings["ExcelName"];
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.Replace("\\bin\\Debug\\", "").Replace("\\bin\\Release\\", "");
            string destinationFolderPath = Path.Combine(basePath, "UploadedExcel", excelName);
            
                // Add columns to DataTable
                foreach (string header in headers)
                {
                    matchedDt.Columns.Add(header);
                    notMatchedDt.Columns.Add(header);
                }
               
            List<string> listOfPDFHavingInvoiceAndPrsentInExcel = new List<string>();
            foreach (var invoice in listOfInvoice)
            {   
                var grandValue = invoiceGrandValue[invoice];
                ReadExcelFile(destinationFolderPath, matchedDt, invoice, grandValue, listOfPDFHavingInvoiceAndPrsentInExcel);
            }

            // Add columns to DataTable
            foreach (string header in headerOfNotFoundInvoiceInExcel)
            {
                invoiceButNotMatchedDt.Columns.Add(header);
            }

            foreach (var invoice in listOfInvoice)
            {
                if(!listOfPDFHavingInvoiceAndPrsentInExcel.Contains(invoice))
                {
                    var grandValue = invoiceGrandValue[invoice];
                    var pdfValue = invoicePDFName[invoice];
                    DataRow dr = invoiceButNotMatchedDt.NewRow();
                    dr[0] = pdfValue;
                    dr[1] = invoice;
                    dr[2] = grandValue;
                    invoiceButNotMatchedDt.Rows.Add(dr);
                } 
            }


            NotMatchedExcelFile(destinationFolderPath, notMatchedDt, listOfInvoice);
            return new PdfExtractionResult
            {
                MatchedTable = matchedDt,
                NotMatchedTable = notMatchedDt,
                ListOfPDFNotHavingInvoice = listOfPDFNotHavingInvoice,
                InvoiceNotMatchedTable = invoiceButNotMatchedDt,
            };
        }

        public static string ExtractTopTextFromPdf(string pdfPath)
        {
            using (var pdfDocument = new PdfDocument(new iText.Kernel.Pdf.PdfReader(pdfPath)))
            {
                var page = pdfDocument.GetPage(1);  // Assuming the text is on the first page
                var strategy = new SimpleTextExtractionStrategy();
                var pageText = PdfTextExtractor.GetTextFromPage(page, strategy);

                var lines = pageText.Split('\n');
                return lines[0];  // Assuming the top text is on the first line
            }
        }
        public static string ExtractGrandTotalFromPdf(string pdfPath, string textTofind)
        {
            using (var pdfDocument = new PdfDocument(new iText.Kernel.Pdf.PdfReader(pdfPath)))
            {
                for (int pageNum = 1; pageNum <= pdfDocument.GetNumberOfPages(); pageNum++)
                {
                    var page = pdfDocument.GetPage(pageNum);
                    var strategy = new SimpleTextExtractionStrategy();
                    var pageText = PdfTextExtractor.GetTextFromPage(page, strategy);

                    var lines = pageText.Split('\n');
                    for (int i = 0; i < lines.Length; i++)
                    {
                        var line = lines[i];
                        if (line.Contains(textTofind))
                        {
                            if (i < lines.Length)
                            {
                                int k = i;
                                if (line.Contains("Bill To") || line.Contains("Authorised Signatory"))
                                {
                                    k++;
                                    if(k < lines.Length)
                                    {
                                        return lines[k];
                                    }else
                                    {
                                        return "";
                                    }
                                    
                                }
                                var totalAmount = ExtractNumberFromLine(lines[k]);
                                if (!string.IsNullOrWhiteSpace(totalAmount))
                                {
                                    return totalAmount;
                                }
                            }
                        }
                    }
                }
            }

            return "";
        }

        public static string ExtractNumberFromLine(string input)
        {
            var values = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return values.Length > 0 ? values.Last() : string.Empty;
        }

        public static void ReadExcelFile(string filePath, DataTable dt, string invoiceNumber, string grandValue, List<string> listOfPDFHavingInvoiceAndPrsentInExcel)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
               
                // Add rows to DataTable 
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    var entityValue = worksheet.Cells[rowNumber, 1].Text;
                    var cscValue = worksheet.Cells[rowNumber, 2].Text;
                    var businessValue = worksheet.Cells[rowNumber, 3].Text;
                    var invoiceValue = worksheet.Cells[rowNumber, 6].Text;
                    var amountINRValue = worksheet.Cells[rowNumber, 11].Text;
                    var amountUSDValue = worksheet.Cells[rowNumber, 12].Text;
                    
                    if (!string.IsNullOrEmpty(invoiceValue) && invoiceValue == invoiceNumber) {
                        DataRow dr = dt.NewRow();
                        var values = new object[]
                           {
                                entityValue,
                                cscValue,
                                businessValue,
                                invoiceValue,
                                amountINRValue,
                                amountUSDValue,
                                grandValue,
                                ParseNumber(amountINRValue) == ParseNumber(grandValue) ? "Matched" : "Not Matched"
                           };

                        for (int i = 0; i < values.Length; i++)
                        {
                            dr[i] = values[i];
                        }

                        dt.Rows.Add(dr);
                        listOfPDFHavingInvoiceAndPrsentInExcel.Add(invoiceValue);
                    } 
                    
                } 
            } 
        }

        public static void CompareInvoiceInExcelFile(string filePath, DataTable dt, PdfExtractedValue pdfExtractedValue, List<string> listOfPDFHavingInvoiceAndPrsentInExcel)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int headerRowNumber = 1;

                // Loop through each row of the Excel sheet starting from row 2
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    Dictionary<string, string> columnValues = new Dictionary<string, string>();

                    var invoiceValue = worksheet.Cells[rowNumber, 6].Text;

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string columnName = worksheet.Cells[headerRowNumber, col].Text;
                        string cellValue = worksheet.Cells[rowNumber, col].Text;

                        if (!string.IsNullOrWhiteSpace(columnName))
                            columnValues[columnName] = cellValue;

                        // Add extra PDF-extracted values
                        switch (columnName)
                        {
                            case "Invoice Date":
                                columnValues["Invoice Date PDF"] = pdfExtractedValue.InvoiceDateValue;
                                break;
                            case "Sum of TOTAL INV AMOUNT INR":
                                columnValues["Grand Value PDF"] = pdfExtractedValue.GrandTotalValue;
                                break;
                            case "Entity Name":
                                columnValues["Bill to Details PDF"] = pdfExtractedValue.BillToDetailsValue;
                                break;
                            case "FX RATE":
                                columnValues["FX RATE PDF"] = pdfExtractedValue.ExchangeRateValue;
                                break;
                            case "Services Month":
                                columnValues["Services Month PDF"] = pdfExtractedValue.ServiceDateValue;
                                break;
                        }
                    }

                    CheckFullyMatchedData(columnValues);

                    // Only add row if invoiceValue matches
                    if (!string.IsNullOrEmpty(invoiceValue) && invoiceValue == pdfExtractedValue.InvoiceValue)
                    {
                        // Ensure DataTable has columns before adding row
                        foreach (var entry in columnValues)
                        {
                            if (!dt.Columns.Contains(entry.Key))
                                dt.Columns.Add(entry.Key, typeof(string));
                        }

                        DataRow dr = dt.NewRow();

                        // Assign values by column name (safer than index)
                        foreach (var entry in columnValues)
                        {
                            dr[entry.Key] = entry.Value;
                        }

                        dt.Rows.Add(dr);
                    }
                }
            }
        }


        //public static void CompareInvoiceInExcelFile(string filePath, DataTable dt, PdfExtractedValue pdfExtractedValue, List<string> listOfPDFHavingInvoiceAndPrsentInExcel)
        //{
        //    using (var package = new ExcelPackage(new FileInfo(filePath)))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        //        int headerRowNumber = 1;

        //        // Add rows to DataTable 
        //        for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
        //        {
        //            var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
        //            Dictionary<string, string> columnValues = new Dictionary<string, string>();
        //            // Loop through each column in the header row to get the column names
        //            var invoiceValue = worksheet.Cells[rowNumber, 6].Text;
        //            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        //            {
        //                // Get the header value (column name) from the first row
        //                string columnName = worksheet.Cells[headerRowNumber, col].Text; 
        //                // Get the value from the current row (rowNumber) and the same column
        //                string cellValue = worksheet.Cells[rowNumber, col].Text;

        //                // Use the column name or index as the key in the dictionary
        //                // If columnName is empty, you could use "Column" + col as a fallback
        //                if (!string.IsNullOrWhiteSpace(columnName))
        //                {
        //                    columnValues.Add(columnName, cellValue);
        //                }
        //                // Add the column name (or index) and cell value to the dictionary

        //                switch (columnName)
        //                {
        //                    case "Invoice Date":
        //                        columnValues.Add("Invoice Date PDF", pdfExtractedValue.InvoiceDateValue);
        //                       // string isInvoiceDateMatched = cellValue == ConvertDate(pdfExtractedValue.InvoiceDateValue) ? "Matched" : "Not Matched";
        //                       // columnValues.Add("Is Invoice Date Matched", isInvoiceDateMatched);
        //                        break;
        //                    case "Sum of TOTAL INV AMOUNT INR":
        //                        columnValues.Add("Grand Value PDF", pdfExtractedValue.GrandTotalValue);
        //                       // string isMatched = ParseNumber(cellValue) == ParseNumber(pdfExtractedValue.GrandTotalValue) ? "Matched" : "Not Matched";
        //                       // columnValues.Add("Is Grand Value Match to INV Amount", isMatched);
        //                        break;
        //                    case "Entity Name": 
        //                         columnValues.Add("Bill to Details PDF", pdfExtractedValue.BillToDetailsValue);
        //                        break;
        //                    case "FX RATE":
        //                        columnValues.Add("FX RATE PDF", pdfExtractedValue.ExchangeRateValue);
        //                        break;
        //                    default:
        //                        break;
        //                } 
        //            }
        //            CheckFullyMatchedData(columnValues); 
        //            if (!string.IsNullOrEmpty(invoiceValue) && invoiceValue == pdfExtractedValue.InvoiceValue)
        //            {
        //                DataRow dr = dt.NewRow();
        //                int index = 0;
        //                foreach (KeyValuePair<string, string> entry in columnValues)
        //                {
        //                    dr[index] = entry.Value;
        //                    index++;
        //                } 
        //                dt.Rows.Add(dr); 
        //            }

        //        }
        //    }
        //}

        public static bool AreDatesSame(string date1, string date2)
        {
            string[] formats = { "dd-MMM-yy", "dd-MM-yyyy" };

            // Parse both dates using the possible formats
            if (DateTime.TryParseExact(date1, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate1) &&
                DateTime.TryParseExact(date2, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate2))
            {
                // Compare the parsed dates
                return parsedDate1.Date == parsedDate2.Date;
            }

            // If either date couldn't be parsed, return false
            return false;
        }
        //static void CheckFullyMatchedData(Dictionary<string, string> columnValues)
        //{
        //    bool isFullyMatched = true;
        //    string notes = string.Empty;
        //    if (columnValues["Entity Name"] != columnValues["Bill to Details PDF"])
        //    {
        //        isFullyMatched = false;
        //        notes = notes + "Entity Name, ";
        //    }

        //    if(ParseNumber(columnValues["Sum of TOTAL INV AMOUNT INR"]) != ParseNumber(columnValues["Grand Value PDF"])) 
        //    { 
        //        isFullyMatched = false;
        //        notes = notes + "Sum of TOTAL INV AMOUNT INR, ";
        //    }

        //    if(columnValues["FX RATE"] != columnValues["FX RATE PDF"])
        //    {
        //        isFullyMatched = false;
        //        notes = notes + "FX RATE, "; 
        //    }

        //    if (AreDatesSame(columnValues["Invoice Date"] ,columnValues["Invoice Date PDF"]) == false )
        //    {
        //        isFullyMatched = false;
        //        notes = notes + "Invoice Date ";
        //    }
        //    if (isFullyMatched) {
        //        columnValues.Add("Matched",  "Fully Matched");
        //    }
        //    else
        //    {
        //        columnValues.Add("Matched", "Partial Matched");
        //    }
        //    if (!string.IsNullOrEmpty(notes)) {
        //        notes = notes + "did not match"; 
        //        columnValues.Add("Notes", notes);
        //    }
        //    else
        //    {
        //        columnValues.Add("Notes",  "Okay");
        //    }
        //}

        private static string GetValue(
    Dictionary<string, string> dict,
    string key)
        {
            return dict.TryGetValue(key, out var value)
                ? value?.Trim() ?? string.Empty
                : string.Empty;
        }


        static void CheckFullyMatchedData(Dictionary<string, string> columnValues)
        {
            bool isFullyMatched = true;
            List<string> mismatchNotes = new List<string>();

            string entityName = GetValue(columnValues, "Entity Name");
            string billToPdf = GetValue(columnValues, "Bill to Details PDF");

            string totalInvAmount = GetValue(columnValues, "Sum of TOTAL INV AMOUNT INR");
            string grandValuePdf = GetValue(columnValues, "Grand Value PDF");

            string fxRate = GetValue(columnValues, "FX RATE");
            string fxRatePdf = GetValue(columnValues, "FX RATE PDF");

            string invoiceDate = GetValue(columnValues, "Invoice Date");
            string invoiceDatePdf = GetValue(columnValues, "Invoice Date PDF");

            string servicesMonths = GetValue(columnValues, "Services Month");
            string servicesMonthsPdf = GetValue(columnValues, "Services Month PDF");

            // Entity name comparison
            if (!string.Equals(entityName, billToPdf, StringComparison.OrdinalIgnoreCase))
            {
                isFullyMatched = false;
                mismatchNotes.Add("Entity Name");
            }

            // Amount comparison
            if (ParseNumber(totalInvAmount) != ParseNumber(grandValuePdf))
            {
                isFullyMatched = false;
                mismatchNotes.Add("Total Amount");
            }

            // FX Rate comparison
            if (!string.Equals(fxRate, fxRatePdf, StringComparison.OrdinalIgnoreCase))
            {
                isFullyMatched = false;
                mismatchNotes.Add("FX Rate");
            }

            // Date comparison
            if (!AreDatesSame(invoiceDate, invoiceDatePdf))
            {
                isFullyMatched = false;
                mismatchNotes.Add("Invoice Date");
            } 
            if (!AreDatesSame(servicesMonths, servicesMonthsPdf))
            {
                isFullyMatched = false;
                mismatchNotes.Add("Services Month");
            }

            // Add / Update result columns safely
            columnValues["Matched"] = isFullyMatched
                ? "Fully Matched"
                : "Partial Matched";

            columnValues["Notes"] = mismatchNotes.Any()
                ? string.Join(", ", mismatchNotes) + " did not match"
                : "Okay";
        }


        static double ParseNumber(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) { return 0; }
            // Remove commas and convert to double
            return double.Parse(value, NumberStyles.AllowThousands | NumberStyles.Float, CultureInfo.InvariantCulture);
        }
        public static void NotMatchedExcelFile(string filePath, DataTable dt, List<string> invoiceNumberList)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    var entityValue = worksheet.Cells[rowNumber, 1].Text;
                    var cscValue = worksheet.Cells[rowNumber, 2].Text;
                    var businessValue = worksheet.Cells[rowNumber, 3].Text;
                    var invoiceValue = worksheet.Cells[rowNumber, 6].Text;
                    var amountINRValue = worksheet.Cells[rowNumber, 11].Text;
                    var amountUSDValue = worksheet.Cells[rowNumber, 12].Text;
                    DataRow dr = dt.NewRow();
                    if (!string.IsNullOrEmpty(invoiceValue) && !invoiceNumberList.Contains(invoiceValue))
                    {

                        var values = new object[]
                        {
                            entityValue,
                            cscValue,
                            businessValue,
                            invoiceValue,
                            amountINRValue,
                            amountUSDValue,
                             "NA",
                            "NA"
                        }; 
                        for (int i = 0; i < values.Length; i++)
                        {
                            dr[i] = values[i];
                        }
                        dt.Rows.Add(dr);
                    }

                }
            }
        }

        public static void NotMatchedTupleIncExcelFile(string filePath, DataTable dt, List<string> invoiceNumberList)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                // Ensure that the DataTable has the correct columns
                if (dt.Columns.Count == 0)
                {
                    // Create columns based on the number of columns in the Excel file
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        // Get the column name from the first row of the worksheet
                        string columnName = worksheet.Cells[1, col].Text;
                        // If the cell is empty, generate a default column name

                        dt.Columns.Add(columnName);
                    }
                }

                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                     
                    var invoiceValue = worksheet.Cells[rowNumber, 6].Text;
                    
                    DataRow dr = dt.NewRow();
                    if (!string.IsNullOrEmpty(invoiceValue) && !invoiceNumberList.Contains(invoiceValue))
                    {
                        for (int colNumber = 1; colNumber <= worksheet.Dimension.End.Column; colNumber++)
                        {
                            dr[colNumber - 1] = worksheet.Cells[rowNumber, colNumber].Text;
                        }

                        dt.Rows.Add(dr);
                    }

                }
            } 
        }
        public static DataTable ReadExcelFileAndFetchData(string filePath, DataTable dt)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                // Ensure that the DataTable has the correct columns
                if (dt.Columns.Count == 0)
                {
                    // Create columns based on the number of columns in the Excel file
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        // Get the column name from the first row of the worksheet
                        string columnName = worksheet.Cells[1, col].Text;
                        // If the cell is empty, generate a default column name
                        
                        dt.Columns.Add(columnName);
                    }
                }
                // Add rows to DataTable 
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    bool isEmpty = CheckIsRowEmpty(row);

                    if (!isEmpty) {
                        DataRow dr = dt.NewRow();
                        // Populate the DataRow with values from the Excel row
                        for (int colNumber = 1; colNumber <= worksheet.Dimension.End.Column; colNumber++)
                        {
                            dr[colNumber - 1] = worksheet.Cells[rowNumber, colNumber].Text;
                        }
                        dt.Rows.Add(dr);
                    }

                     
                }
            }
            return dt;
        }

        public static bool CheckIsRowEmpty(ExcelRange  row)
        {
            bool isRowEmpty = true; 

            foreach (var cell in row)
            {
                // Check if the cell is null or empty
                if (!string.IsNullOrWhiteSpace(cell.Text))
                {
                    isRowEmpty = false;
                    break;
                }
            }
            return isRowEmpty;
        }
        public static PdfExtractionResult ExtractICSAAndAddress( )
        {
              string excelFilePath = GetUploadedExcelPath();
            DataTable dataTable = new DataTable();
            // Add columns to DataTable
            foreach (string header in headerOfICSAVaryingAddress)
            {
                dataTable.Columns.Add(header);
            }
            var icsaAddressPairs = new Dictionary<string, string>();
            // Create a dictionary to store ICSA codes and their corresponding addresses
            Dictionary<string, ExcelICSAData> icsaAddresses = new Dictionary<string, ExcelICSAData>();


            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 

               
                ExcelICSAData excelICSAData = new ExcelICSAData();
                // Add rows to DataTable 
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    Dictionary<string, string> columnValues = new Dictionary<string, string>();
                    // Loop through each column in the header row to get the column names
                    var icsaValue =  worksheet.Cells[rowNumber, 15].Text;
                    var addressValue = worksheet.Cells[rowNumber, 13].Text;
                    var invoiceValue = worksheet.Cells[rowNumber, 6].Text;
                    if (!string.IsNullOrEmpty(icsaValue) && !string.IsNullOrEmpty(addressValue) )
                    {
                        if (icsaAddresses.ContainsKey(icsaValue))
                        {
                            if (icsaAddresses[icsaValue] != null) {
                                if (icsaAddresses[icsaValue].Addresses == null) {
                                    icsaAddresses[icsaValue].Addresses = new List<string>();
                                    icsaAddresses[icsaValue].InvoiceValues = new List<string>();
                                    icsaAddresses[icsaValue].Addresses.Add(addressValue);
                                    icsaAddresses[icsaValue].InvoiceValues.Add(invoiceValue);
                                }
                                else
                                {
                                    if (!icsaAddresses[icsaValue].Addresses.Contains(addressValue))
                                    {
                                        icsaAddresses[icsaValue].Addresses.Add(addressValue);
                                        icsaAddresses[icsaValue].InvoiceValues.Add(invoiceValue);
                                    }

                                }

                            }
                            
                        }
                        else
                        {
                            icsaAddresses[icsaValue] = new ExcelICSAData();
                            icsaAddresses[icsaValue].Addresses = new List<string> { addressValue };
                            icsaAddresses[icsaValue].InvoiceValues = new List<string> {invoiceValue};
                        }
                    } 
                }
                
            }
            foreach (KeyValuePair<string, ExcelICSAData> kvp in icsaAddresses)
            {
                if (kvp.Value != null) {
                   
                    if (kvp.Value.Addresses.Count >1) {
                        for (int i = 0; i < kvp.Value.Addresses.Count; i++)
                        {
                            DataRow dr = dataTable.NewRow();
                            dr["INVOICE NUMBER"] = kvp.Value.InvoiceValues[i];
                            dr["ICSA"] = kvp.Key;
                            dr["Address"] = kvp.Value.Addresses[i];
                            dataTable.Rows.Add(dr);
                        }
                    }
                } 
            }
            return new PdfExtractionResult
            {
                ICSAVaryingAddress = dataTable,
            }; 
        }
    }
}
