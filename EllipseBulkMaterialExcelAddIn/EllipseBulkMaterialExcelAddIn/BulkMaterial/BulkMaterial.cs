using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace EllipseBulkMaterialExcelAddIn
{
    public class BulkMaterialItem
    {
        [CsvColumn(FieldIndex = 1)]
        public string WarehouseId { get; set; }

        [CsvColumn(FieldIndex = 2, OutputFormat = "yyyyMMdd")]
        public string DefaultUsageDate { get; set; }

        [CsvColumn(FieldIndex = 3)]
        public string UserId { get; set; }

        [CsvColumn(FieldIndex = 4)]
        public string EquipmentReference { get; set; }

        [CsvColumn(FieldIndex = 5)]
        public string BulkMaterialTypeId { get; set; }

        [CsvColumn(FieldIndex = 6)]
        public string Quantity { get; set; }
    }
}
