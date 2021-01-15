using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace EllipseMSO261ProntoPagoExcelAddIn
{
    public class InvoiceParameters
    {
        [CsvColumn(FieldIndex = 1)]
        public double Percentage { get; set; }

        [CsvColumn(FieldIndex = 2)]
        public double Days { get; set; }

        [CsvColumn(FieldIndex = 3)]
        public string Branchcode { get; set; }

        [CsvColumn(FieldIndex = 4)]
        public string Bankaccount { get; set; }
    }
}
