using System.Diagnostics.CodeAnalysis;
using EllipseWorkRequestClassLibrary.WorkRequestService;

namespace EllipseWorkRequestClassLibrary
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class WorkRequest
    {
        public string activityClass;
        public string assignPerson;
        public string classification;
        public string classificationDescription;
        public string closedBy;
        public string closedDate;
        public string closedTime;
        public string contactId;
        public string copyRequestId;
        public string custPOItemNoRef;
        public string custPONoRef;
        public string customerNo;
        public string employee;
        public string equipmentNo;
        public string equipmentRef;
        public string estimateNo;
        public string location;
        public string ownerId;
        public string priorityCode;
        public string priorityCodeDescription;
        public string programCode;
        public string raisedDate;
        public string raisedTime;
        public string raisedUser;
        public string region;
        public string regionDescription;
        public string requestId;
        public string requestIdDescription1;
        public string requestIdDescription2;
        public string requestType;
        public string requestTypeDescription;
        public string requestorId;
        public string requiredByDate;
        public string requiredByTime;
        public string riskCode1;
        public string riskCode10;
        public string riskCode2;
        public string riskCode3;
        public string riskCode4;
        public string riskCode5;
        public string riskCode6;
        public string riskCode7;
        public string riskCode8;
        public string riskCode9;
        public string source;
        public string sourceDescription;
        public string sourceReference;
        public string standardJob;
        public string standardJobDistrict;
        public string userStatus;
        public string userStatusDescription;
        public string workGroup;
        public string status;
        public decimal priorityValue;
        public bool priorityValueFieldSpecified;
        public ServiceLevelAgreement ServiceLevelAgreement;
        private ExtendedDescription _extendedDescription;

        public WorkRequest()
        {
            ServiceLevelAgreement = new ServiceLevelAgreement();
        }
        public ExtendedDescription GetExtendedDescription(string urlService, OperationContext opContext)
        {
            if (_extendedDescription != null) return _extendedDescription;

            _extendedDescription = WorkRequestActions.GetWorkRequestExtendedDescription(urlService, opContext, requestId);

            return _extendedDescription;
        }

        public void SetExtendedDescription(string header, string body)
        {
            if (_extendedDescription == null)
                _extendedDescription = new ExtendedDescription();
            _extendedDescription.Header = header;
            _extendedDescription.Body = body;
        }
    }



    
}
