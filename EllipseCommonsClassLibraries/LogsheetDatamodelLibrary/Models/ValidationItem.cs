using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using LogsheetDatamodelLibrary.Controllers;

namespace LogsheetDatamodelLibrary.Models
{
    public class ValidationItem : ISimpleObjectModelSql
    {
        private string _sourceName; //
        private ValidationSource _validationSource;

        public int? Id;
        public string Description;
        public string SourceTable;
        public string SourceColumn;
        public bool Sortable;
        public bool DistinctFilter;
        private bool _autoPull = LsdmConfig.AutoPullValues;

        public void SetFromDataRecord(IDataRecord dr)
        {
            SourceName = dr["source_name"].ToString().Trim();
            Id = MyUtilities.ToIntegerNull(dr["valid_item_id"].ToString().Trim());
            Description = dr["valid_item_desc"].ToString().Trim();
            SourceTable = dr["valid_item_source_table"].ToString().Trim();
            SourceColumn = dr["valid_item_source_column"].ToString().Trim();
            Sortable = MyUtilities.IsTrue(dr["sortable"].ToString().Trim());
            DistinctFilter = MyUtilities.IsTrue(dr["distinct_filter"].ToString().Trim());
        }

        public ValidationItem()
        {

        }

        public ValidationItem(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }

        public string SourceName
        {
            get => _sourceName;
            set
            {
                _validationSource = null;
                _sourceName = value;
            }
        }
        public ValidationSource ValidationSource
        {
            get
            {
                if (_validationSource != null || !_autoPull)
                    return _validationSource;

                _validationSource = ValidationSourceController.ReadFirst(_sourceName);
                if (_validationSource != null)
                    _sourceName = _validationSource.Name;
                return _validationSource;
            }
            set
            {
                _validationSource = value;
                _sourceName = _validationSource?.Name;
            }
        }
        public bool Equals(ValidationItem validationItem)
        {
            if (Id != validationItem.Id)
                return false;
            if (!MyUtilities.EqualUpper(Description, validationItem.Description))
                return false;
            if (!MyUtilities.EqualUpper(SourceTable, validationItem.SourceTable))
                return false;
            if (!MyUtilities.EqualUpper(SourceColumn, validationItem.SourceColumn))
                return false;
            if (!Sortable == validationItem.Sortable)
                return false;
            if (!DistinctFilter == validationItem.DistinctFilter)
                return false;
            return true;
        }
        public ValidationItem(string sourceName, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {
            _sourceName = sourceName;
            Id = id;
            Description = description;
            SourceTable = sourceTable;
            SourceColumn = sourceColumn;
            Sortable = sortable;
            DistinctFilter = distinctFilter;
            
        }

        public ValidationItem(ValidationSource validationSource, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {
            _validationSource = validationSource;

            Id = id;
            Description = description;
            SourceTable = sourceTable;
            SourceColumn = sourceColumn;
            Sortable = sortable;
            DistinctFilter = distinctFilter;
        }

        public ValidationItem(ValidationItem validationItem)
        {
            if(validationItem.ValidationSource != null)
                _validationSource = validationItem.ValidationSource;
            else
                _sourceName = validationItem._sourceName;

            Id = validationItem.Id;
            Description = validationItem.Description;
            SourceTable = validationItem.SourceTable;
            SourceColumn = validationItem.SourceColumn;
            Sortable = validationItem.Sortable;
            DistinctFilter = validationItem.DistinctFilter;
        }

        public void SetAutoPull(bool autoPull)
        {
            _autoPull = autoPull;
        }
    }
}
