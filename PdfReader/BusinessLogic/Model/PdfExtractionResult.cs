using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfReader.BusinessLogic.Model
{
    public class PdfExtractionResult
    {
        public DataTable MatchedTable { get; set; }
        public DataTable NotMatchedTable { get; set; }

        public DataTable InvoiceNotMatchedTable { get; set; }
        public List<string> ListOfPDFNotHavingInvoice { get; set; }
        public List<string> ListOfPDFHavingInvoiceNotPrsentInExcel { get; set; }
        public DataTable SourceExcelTable { get; set; }

        public DataTable ICSAVaryingAddress { get; set; }
    }

    public class ExcelICSAData
    {
        public List<string> InvoiceValues { get; set; }
        public List<string> Addresses { get; set; }
    }
    public class PdfExtractedValue
    {
        public string InvoiceValue { get; set; }
        public string GrandTotalValue { get; set; }
        public string InvoiceDateValue { get; set; }
        public string BillToDetailsValue { get;set; }
        public string ExchangeRateValue {  get; set; }
    }
}
