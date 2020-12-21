using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using LogsheetDatamodelLibrary.Configuration;
using LogsheetDatamodelLibrary.Controllers;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class ModelAttribute : ISimpleObjectModelSql
    {
        private string _modelId; //
        private Datamodel _datamodel;
        public string Id;
        public string Description;
        public string DataType;
        public int? SheetIndex;
        public int? MaxLength;
        public int? MaxPrecision;
        public int? MaxScale;
        public bool AllowNull;
        public string DefaultValue;
        public bool ActiveStatus;

        private int? _measureId; //
        private Measure _measure;
        private int? _validationItemId; //
        private ValidationItem _validationItem;
        private bool _autoPull = LsdmConfig.AutoPullValues;
        public string ModelId
        {
            get => _modelId;
            set
            {
                _datamodel = null;
                _modelId = value;
            }
        }
        public Datamodel Datamodel
        {
            get
            {
                if (_datamodel != null || !_autoPull)
                    return _datamodel;

                _datamodel = DatamodelController.ReadFirst(_modelId);
                if (_datamodel != null)
                    _modelId = _datamodel.Id;
                return _datamodel;
            }
            set
            {
                _datamodel = value;
                _modelId = _datamodel?.Id;
            }
        }
        public int? MeasureId
        {
            get => _measureId;
            set
            {
                _measure = null;
                _measureId = value;
            }
        }
        public Measure Measure
        {
            get
            {
                if (_measure != null || !_autoPull)
                    return _measure;

                _measure = MeasureController.ReadFirst(_measureId);
                if (_measure != null)
                    _measureId = _measure.Id;
                return _measure;
            }
            set
            {
                _measure = value;
                _measureId = _measure?.Id;
            }
        }
        public int? ValidationItemId
        {
            get => _validationItemId;
            set
            {
                _validationItem = null;
                _validationItemId = value;
            }
        }

        public void SetFromDataRecord(IDataRecord dr)
        {
            ModelId = dr["model_id"].ToString().Trim();
            Id = dr["attribute_id"].ToString().Trim();
            Description = dr["attribute_desc"].ToString().Trim();
            DataType = dr["datatype"].ToString().Trim();
            SheetIndex = MyUtilities.ToIntegerNull(dr["sheet_index"].ToString().Trim());
            MaxLength = MyUtilities.ToIntegerNull(dr["max_length"].ToString().Trim());
            MaxPrecision = MyUtilities.ToIntegerNull(dr["max_precision"].ToString().Trim());
            MaxScale = MyUtilities.ToIntegerNull(dr["max_scale"].ToString().Trim());
            AllowNull = MyUtilities.IsTrue(dr["allow_null"].ToString().Trim());
            DefaultValue = dr["default_value"].ToString().Trim();
            ActiveStatus = MyUtilities.IsTrue(dr["active_status"].ToString().Trim());
            MeasureId = MyUtilities.ToIntegerNull(dr["measure_id"].ToString().Trim());
            ValidationItemId = MyUtilities.ToIntegerNull(dr["valid_item_id"].ToString().Trim());
        }

        public ModelAttribute()
        {

        }

        public ModelAttribute(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }
        public ValidationItem ValidationItem
        {
            get
            {
                if (_validationItem != null || !_autoPull)
                    return _validationItem;

                _validationItem = ValidationItemController.Read(_validationItemId);
                if (_validationItem != null)
                    _validationItemId = _validationItem.Id;
                return _validationItem;
            }
            set
            {
                _validationItem = value;
                _validationItemId = _validationItem?.Id;
            }
        }
        public bool Equals(ModelAttribute modelAttribute)
        {
            if (!MyUtilities.EqualUpper(Id, modelAttribute.Id))
                return false;
            if (!MyUtilities.EqualUpper(Description, modelAttribute.Description))
                return false;
            if (!MyUtilities.EqualUpper(DataType, modelAttribute.DataType))
                return false;
            if (SheetIndex != modelAttribute.SheetIndex)
                return false;
            if (MaxLength != modelAttribute.MaxLength)
                return false;
            if (MaxPrecision != modelAttribute.MaxPrecision)
                return false;
            if (MaxScale != modelAttribute.MaxScale)
                return false;
            if (AllowNull != modelAttribute.AllowNull)
                return false;
            if (!MyUtilities.EqualUpper(DefaultValue, modelAttribute.DefaultValue))
                return false;
            if (ActiveStatus != modelAttribute.ActiveStatus)
                return false;
            return true;
        }

        public ModelAttribute(string modelId, string id, string description, string dataType, int? sheetIndex, int? maxLength,  int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, int? measureId, int? validItemId)
        {
            _modelId = modelId;
            Id = id;
            Description = description;
            DataType = dataType;
            SheetIndex = sheetIndex;
            MaxLength = maxLength;
            MaxPrecision = maxPrecision;
            MaxScale = maxScale;
            DefaultValue = defaultValue;
            AllowNull = allowNull;
            ActiveStatus = activeStatus;
            _measureId = measureId;
            _validationItemId = validItemId;
        }

        public ModelAttribute(Datamodel datamodel, string id, string description, string dataType, int? sheetIndex, int? maxLength, int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, Measure measure, ValidationItem validItem)
        {
            Id = id;
            Description = description;
            DataType = dataType;
            SheetIndex = sheetIndex;
            MaxLength = maxLength;
            MaxPrecision = maxPrecision;
            MaxScale = maxScale;
            DefaultValue = defaultValue;
            AllowNull = allowNull;
            ActiveStatus = activeStatus;

            _datamodel = datamodel;
            _measure = measure;
            _validationItem = validItem;
        }

        public ModelAttribute(ModelAttribute modelAttribute)
        {
            Id = modelAttribute.Id;
            Description = modelAttribute.Description;
            DataType = modelAttribute.DataType;
            SheetIndex = modelAttribute.SheetIndex;
            MaxLength = modelAttribute.MaxLength;
            MaxPrecision = modelAttribute.MaxPrecision;
            MaxScale = modelAttribute.MaxScale;
            DefaultValue = modelAttribute.DefaultValue;
            AllowNull = modelAttribute.AllowNull;
            ActiveStatus = modelAttribute.ActiveStatus;

            if (modelAttribute.Datamodel != null)
                _datamodel = modelAttribute.Datamodel;
            else
                _modelId = modelAttribute.ModelId;

            if (modelAttribute.Measure != null)
                _measure = modelAttribute.Measure;
            else
                _measureId = modelAttribute.MeasureId;

            if (modelAttribute.ValidationItem != null)
                _validationItem = modelAttribute.ValidationItem;
            else
                _validationItemId = modelAttribute.ValidationItemId;
        }
        public void SetAutoPull(bool autoPull)
        {
            _autoPull = autoPull;
        }
    }

}
