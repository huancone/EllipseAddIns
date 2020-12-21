using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using LogsheetDatamodelLibrary.Configuration;
using LogsheetDatamodelLibrary.Controllers;
using SharedClassLibrary;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class Datamodel : ISimpleObjectModelSql
    {
        public string Id;
        public string Description;
        public DateTime CreationDate;
        public string CreationUser;
        public DateTime LastModDate;
        public string LastModUser;
        public bool ActiveStatus;
        private List<ModelAttribute> _modelAttributes;
        private bool _autoPull = LsdmConfig.AutoPullValues;

        public void SetFromDataRecord(IDataRecord dr)
        {
            Id = dr["model_id"].ToString().Trim();
            Description = dr["model_desc"].ToString().Trim();
            CreationDate = Convert.ToDateTime(dr["creation_date"]);
            CreationUser = dr["creation_user"].ToString().Trim();
            LastModDate = Convert.ToDateTime(dr["last_mod_date"]);
            LastModUser = dr["last_mod_user"].ToString().Trim();
            ActiveStatus = MyUtilities.IsTrue(dr["active_status"].ToString().Trim());
        }

        public Datamodel()
        {

        }

        public Datamodel(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }
        public List<ModelAttribute> ModelAttributes
        {
            get
            {
                if (_modelAttributes == null && _autoPull)
                {
                    _modelAttributes = DatamodelController.GetModelAttributes(Id);
                }
                return _modelAttributes;
            } 
            set => _modelAttributes = value;
        }

        public bool Equals(Datamodel datamodel, bool compareAttributeList = false)
        {
            if (!MyUtilities.EqualUpper(Id, datamodel.Id))
                return false;
            if (!MyUtilities.EqualUpper(Description, datamodel.Description))
                return false;
            if (CreationDate != datamodel.CreationDate)
                return false;
            if (!MyUtilities.EqualUpper(CreationUser, datamodel.CreationUser))
                return false;
            if (LastModDate != datamodel.LastModDate)
                return false;
            if (!MyUtilities.EqualUpper(LastModUser, datamodel.LastModUser))
                return false;
            if (ActiveStatus != datamodel.ActiveStatus)
                return false;

            if (compareAttributeList)
                return EqualsAttributesList(datamodel.ModelAttributes);
            return true;
        }

        private bool EqualsAttributesList(List<ModelAttribute> attributeList)
        {
            var thisEmpty = _modelAttributes == null || _modelAttributes.Count == 0;
            var otherEmpty = attributeList == null || attributeList.Count == 0;

            if (thisEmpty && otherEmpty)
                return true;

            if (_modelAttributes != null && attributeList != null)
            {
                if (_modelAttributes.Count != attributeList.Count)
                    return false;
                foreach (var att1 in _modelAttributes)
                {
                    foreach (var att2 in attributeList)
                    {
                        if (att1.Equals(att2))
                            break;
                        return false;
                    }
                }
            }
            return true;
        }
        public List<ModelAttribute> PullModelAttributes(bool activeOnly = true, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            if (string.IsNullOrWhiteSpace(Id))
                throw new ArgumentNullException(Resources.Error_InvalidId, nameof(Id));
            _modelAttributes = DatamodelController.GetModelAttributes(Id, activeOnly, connector);
            return _modelAttributes;
        }

        public ModelAttribute AddModelAttribute(ref ModelAttribute modelAttribute)
        {
            if(_modelAttributes == null)
                _modelAttributes = new List<ModelAttribute>();

            _modelAttributes.Add(modelAttribute);
            _modelAttributes.Sort((x, y) => MyUtilities.CompareIntNull(x.SheetIndex, y.SheetIndex));
            return modelAttribute;
        }

        public void SetAutoPull(bool autoPull)
        {
            _autoPull = autoPull;
        }

    }
}
