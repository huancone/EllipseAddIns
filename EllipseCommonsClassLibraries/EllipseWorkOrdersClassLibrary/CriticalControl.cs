namespace EllipseWorkOrdersClassLibrary
{
    public class CriticalControl
    {
        public string WorkOrder;
        public string TaskNo;
        public string TaskDescription;
        public string WorkOrderDescription;
        public string CriticalCode;
        public string CriticalDescription;
        public string EquipmentNo;
        public string AssignPerson;
        public string Department;
        public string Quartermaster;
        public string PlanStartDate;
        public string RaisedDate;
        public string MaintSchTask;
        public string StdJobNo;
        public string Status;
        public string CompletedCode;
        public string CompletedBy;
        public string CompletedDate;
        public string InstructionsCode;
        public string InstructionsText;
        public string FrequencyText;
    }
    public class CriticalControlDefaultExport
    {
        public bool WorkOrder = true;
        public bool TaskNo = true;
        public bool TaskDescription = true;
        public bool WorkOrderDescription = true;
        //public bool CriticalCode = false;
        public bool CriticalDescription = true;
        public bool EquipmentNo = false;
        public bool AssignPerson = true;
        public bool Department = true;
        public bool Quartermaster = false;
        public bool PlanStartDate = true;
        public bool RaisedDate = false;
        public bool MaintSchTask = false;
        public bool StdJobNo = true;
        public bool Status = true;
        public bool CompletedCode = false;
        public bool CompletedBy = false;
        public bool CompletedDate = false;
        //public bool InstructionsCode = false;
        public bool InstructionsText = true;
        public bool FrequencyText = true;
    }
}
