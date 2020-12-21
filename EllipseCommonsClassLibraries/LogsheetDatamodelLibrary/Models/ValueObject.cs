using System;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class ValueObject : ISimpleObjectModelSql
    {
        public string ModelId;//
        public int? SheetId;//
        public string AttributeId;//
        public string DataType { get; private set; }

        private string _valueString;
        private decimal? _valueDecimal;
        private DateTime? _valueDateTime;

        public void SetFromDataRecord(IDataRecord dr)
        {
            var modelId = dr["model_id"].ToString().Trim();
            var attributeId = dr["attribute_id"].ToString().Trim();
            var sheetId = MyUtilities.ToIntegerNull(dr["sheet_id"].ToString().Trim());
            var dataType = dr["datatype"].ToString().Trim();

            ModelId = modelId;
            AttributeId = attributeId;
            SheetId = sheetId;


            if (dataType.Equals(DataTypes.Numeric))
            {
                var numericValue = dr.GetDecimal(dr.GetOrdinal("numeric_value"));
                SetValue(numericValue);
            }
            else if (dataType.Equals(DataTypes.Varchar))
            {
                var varcharValue = dr.GetString(dr.GetOrdinal("varchar_value"));
                SetValue(varcharValue);
            }
            else if (dataType.Equals(DataTypes.DateTime) || dataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
            {
                var dateTimevalue = dr.GetDateTime(dr.GetOrdinal("datetime_value"));
                SetValue(dateTimevalue);
            }
            else if (dataType.Equals(DataTypes.Text))
            {
                var textValue = dr.GetString(dr.GetOrdinal("text_value"));
                SetValue(textValue);
            }
        }
        public ValueObject()
        {

        }

        public ValueObject(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }
        public dynamic Value => GetValue();
        public string StringValue => GetValue().ToString();
        public void SetValue(int value)
        {
            _valueString = null;
            _valueDecimal = value;
            _valueDateTime = null;

            if(string.IsNullOrWhiteSpace(DataType))
                DataType = DataTypes.Numeric;
        }
        public void SetValue(float value)
        {
            _valueString = null;
            _valueDecimal = (decimal?) value;
            _valueDateTime = null;

            if (string.IsNullOrWhiteSpace(DataType))
                DataType = DataTypes.Numeric;
        }
        public void SetValue(double value)
        {
            _valueString = null;
            _valueDecimal = (decimal?) value;
            _valueDateTime = null;

            if (string.IsNullOrWhiteSpace(DataType))
                DataType = DataTypes.Numeric;
        }
        public void SetValue(string value)
        {
            _valueString = value;
            _valueDecimal = null;
            _valueDateTime = null;
            if (string.IsNullOrWhiteSpace(DataType))
                DataType = DataTypes.Varchar;
        }
        public void SetValue(decimal value)
        {
            _valueString = null;
            _valueDecimal = value;
            _valueDateTime = null;
            if (string.IsNullOrWhiteSpace(DataType))
                DataType = DataTypes.Numeric;
        }
        public void SetValue(DateTime value)
        {
            _valueString = null;
            _valueDecimal = null;
            _valueDateTime = value;
            if (string.IsNullOrWhiteSpace(DataType))
                DataType = DataTypes.DateTime;
        }

        public void SetValue(object value)
        {
            if (string.IsNullOrWhiteSpace(DataType))
            {
                if (decimal.TryParse(value.ToString(), out _))
                    DataType = DataTypes.Numeric;
            }

            if (DataType.Equals(DataTypes.Numeric))
                SetValue((decimal) value);
            else if (DataType.Equals(DataTypes.Varchar))
                SetValue((string) value);
            else if (DataType.Equals(DataTypes.Text))
                SetValue((string)value);
            else if (DataType.Equals(DataTypes.DateTime) || DataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                SetValue((DateTime)value);
            else
                SetValue((string)value);
        }
        public void SetDataType(string dataType)
        {
            if(!DataTypes.GetList().Contains(dataType))
                throw new ArgumentException(LsdmResource.Error_DataType_Invalid, nameof(dataType));

            DataType = dataType;
        }

        private dynamic GetValue()
        {
            if (_valueString != null)
                return _valueString;

            if (_valueDecimal != null)
                return _valueDecimal;

            if (_valueDateTime != null)
                return _valueDateTime;

            return null;
        }

    }
}
