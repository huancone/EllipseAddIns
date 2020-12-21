using System;
using System.Collections.Generic;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using LogsheetDatamodelLibrary.Controllers;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class Datasheet : ISimpleObjectModelSql
    {
        public int? Id;
        public DateTime Date;
        public string Shift;
        public string SequenceId;
        public string CreationUser;
        public DateTime CreationDate;
        public string LastModUser;
        public DateTime LastModDate;
        private string _modelId;//
        private Datamodel _datamodel;
        private List<ModelAttribute> _headerAttributes;
        private List<ValueObject> _valueObjects;
        private bool _autoPull = LsdmConfig.AutoPullValues;

        public void SetFromDataRecord(IDataRecord dr)
        {
            ModelId = dr["model_id"].ToString().Trim();
            Id = MyUtilities.ToIntegerNull(dr["sheet_id"].ToString().Trim());
            Date = Convert.ToDateTime(dr["sheet_date"]);
            Shift = dr["shift"].ToString().Trim();
            SequenceId = dr["sequence_id"].ToString().Trim();
            CreationDate = Convert.ToDateTime(dr["creation_date"]);
            CreationUser = dr["creation_user"].ToString().Trim();
            LastModDate = Convert.ToDateTime(dr["last_mod_date"]);
            LastModUser = dr["last_mod_user"].ToString().Trim();
        }

        public Datasheet()
        {

        }

        public Datasheet(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }

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


        public List<ModelAttribute> HeaderAttributes
        {
            get
            {
                if (_headerAttributes != null || !_autoPull)
                    return _headerAttributes;
                _headerAttributes = DatamodelController.GetModelAttributes(ModelId);
                return _headerAttributes;
            }
            set => _headerAttributes = value;
        }

        public List<ValueObject> ValueObjects
        {
            get
            {
                if (_valueObjects != null || !_autoPull)
                    return _valueObjects;
                _valueObjects = DatasheetController.GetDataSheetValueObjects(Id);

                return _valueObjects;
            }
            set => _valueObjects = value;
        }
        public Datamodel PullDatamodel(bool activeOnly = true, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            if (string.IsNullOrWhiteSpace(ModelId))
                throw new ArgumentNullException(Resources.Error_InvalidId, nameof(ModelId));
            _datamodel = DatamodelController.ReadFirst(ModelId, connector);
            return _datamodel;
        }
        public List<ValueObject> PullValueObjects(OracleConnector connector = null)
        {
            _valueObjects = DatasheetController.GetDataSheetValueObjects(Id, connector);

            return _valueObjects;
        }
        public ReplyMessage PushValueObjects(OracleConnector connector = null)
        {
            foreach (var item in _valueObjects)
            {
                if (item.SheetId == null)
                    item.SheetId = Id;
                if (item.ModelId == null)
                    item.ModelId = ModelId;
            }
            return ValueObjectController.Create(_valueObjects);
        }
        public void SetAutoPull(bool autoPull)
        {
            _autoPull = autoPull;
        }
    }
}
