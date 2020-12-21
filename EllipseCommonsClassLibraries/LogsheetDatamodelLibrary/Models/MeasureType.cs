using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class MeasureType : ISimpleObjectModelSql
    {
        public int? Id;
        public string Description;

        public void SetFromDataRecord(IDataRecord dr)
        {
            Id = MyUtilities.ToInteger(dr["measure_type_id"].ToString().Trim());
            Description = dr["measure_type_desc"].ToString().Trim();
        }

        public MeasureType()
        {

        }

        public MeasureType(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }
        public MeasureType(int? id, string description)
        {
            Id = id;
            Description = description;
        }
        public bool Equals(MeasureType measureType)
        {
            if (Id != measureType.Id)
                return false;
            if (!MyUtilities.EqualUpper(Description, measureType.Description))
                return false;
            return true;
        }


    }
}
