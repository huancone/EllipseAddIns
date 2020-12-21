using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using LogsheetDatamodelLibrary.Controllers;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class Measure : ISimpleObjectModelSql
    {
        public int? Id;
        public string Code;
        public string Name;
        public string Description;
        public string Units;
        public bool ActiveStatus;
        private int? _measureTypeId; //
        private MeasureType _measureType;
        private bool _autoPull = LsdmConfig.AutoPullValues;

        public void SetFromDataRecord(IDataRecord dr)
        {
            Id = MyUtilities.ToInteger(dr["measure_id"].ToString().Trim());
            Code = dr["measure_code"].ToString().Trim();
            Name = dr["measure_name"].ToString().Trim();
            Description = dr["measure_desc"].ToString().Trim();
            Units = dr["measure_units"].ToString().Trim();
            ActiveStatus = MyUtilities.IsTrue(dr["measure_status"].ToString().Trim());
            _measureTypeId = MyUtilities.ToInteger(dr["measure_type_id"].ToString().Trim());

            try
            {
                var measureTypeId = MyUtilities.ToInteger(dr["measure_type_id"].ToString().Trim());
                var measureTypeDesc = dr["measure_type_desc"].ToString().Trim();
                var measureType = new MeasureType(measureTypeId, measureTypeDesc);
                MeasureType = measureType;
            }
            catch
            {
                //ignored
            }
        }

        public Measure()
        {

        }

        public Measure(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }

        public int? MeasureTypeId
        {
            get => _measureTypeId;
            set
            {
                _measureType = null;
                _measureTypeId = value;
            }
        }

        public MeasureType MeasureType
        {
            get
            {
                if (_measureType != null || !_autoPull)
                    return _measureType;

                _measureType = MeasureTypeController.ReadFirst(_measureTypeId);
                if (_measureType != null)
                    _measureTypeId = _measureType.Id;
                return _measureType;
            }
            set
            {
                _measureType = value;
                _measureTypeId = _measureType?.Id;
            }
        }

        public Measure(int? id, string code, string name, string description, string units, bool activeStatus, int? measureTypeId)
        {
            Id = id;
            Name = name;
            Description = description;
            Code = code;
            Units = units;
            ActiveStatus = activeStatus;
            _measureTypeId = measureTypeId;
        }

        public Measure(int? id, string code, string name, string description, string units, bool activeStatus, MeasureType measureType)
        {
            Id = id;
            Name = name;
            Description = description;
            Code = code;
            Units = units;
            ActiveStatus = activeStatus;
            _measureTypeId = measureType.Id;
            _measureType = measureType;
        }

        public Measure(Measure measure)
        {
            Id = measure.Id;
            Name = measure.Name;
            Description = measure.Description;
            Code = measure.Code;
            Units = measure.Units;
            ActiveStatus = measure.ActiveStatus;

            _measureType = measure._measureType;
            _measureTypeId = _measureType?.Id ?? measure._measureTypeId;

        }

        public bool Equals(Measure measure)
        {
            if (Id != measure.Id)
                return false;
            if (!MyUtilities.EqualUpper(Code, measure.Code))
                return false;
            if (!MyUtilities.EqualUpper(Name, measure.Name))
                return false;
            if (!MyUtilities.EqualUpper(Description, measure.Description))
                return false;
            if (!MyUtilities.EqualUpper(Units, measure.Units))
                return false;
            if (ActiveStatus != measure.ActiveStatus)
                return false;
            if (_measureTypeId != measure._measureTypeId)
                return false;
            return true;
        }

        public void SetAutoPull(bool autoPull)
        {
            _autoPull = autoPull;
        }
    }
}
