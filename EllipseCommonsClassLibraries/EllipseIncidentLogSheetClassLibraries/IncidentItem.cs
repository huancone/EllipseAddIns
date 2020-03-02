using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseIncidentLogSheetClassLibraries
{
    public class IncidentItem
    {
        //Corresponden a un encabezado general
        //public string WorkGroup;
        //public string Date;
        //public string Shift;
        //public string WorkOrderPrefix;

        public string RaisedTime;
        public string IncidentDescription;
        public string MaintenanceType;
        public string Originator;
        public string JobDurationFinish;
        public string IncidentStatus;//C - Closed, O - Open
        public string EquipmentReference;
        public string ComponentCode;
        public string ModifierCode;
        public string JobDurationCode;
        public string DurationHours;
        public string StandardJob;
        public string CorrectiveDescription;
        public string WorkOrder;

        public bool Equals(IncidentItem item, bool ignoreNullValues = true)
        {
            if (ignoreNullValues)
            {
                if ((string.IsNullOrWhiteSpace(item.RaisedTime) || this.RaisedTime == item.RaisedTime) &&
                   (string.IsNullOrWhiteSpace(item.IncidentDescription) || this.IncidentDescription == item.IncidentDescription) &&
                   (string.IsNullOrWhiteSpace(item.MaintenanceType) || this.MaintenanceType == item.MaintenanceType) &&
                   (string.IsNullOrWhiteSpace(item.Originator) || this.Originator == item.Originator) &&
                   (string.IsNullOrWhiteSpace(item.JobDurationFinish) || this.JobDurationFinish == item.JobDurationFinish) &&
                   (string.IsNullOrWhiteSpace(item.IncidentStatus) || this.IncidentStatus == item.IncidentStatus) &&
                   (string.IsNullOrWhiteSpace(item.EquipmentReference) || this.EquipmentReference == item.EquipmentReference) &&
                   (string.IsNullOrWhiteSpace(item.ComponentCode) || this.ComponentCode == item.ComponentCode) &&
                   (string.IsNullOrWhiteSpace(item.ModifierCode) || this.ModifierCode == item.ModifierCode) &&
                   (string.IsNullOrWhiteSpace(item.JobDurationCode) || this.JobDurationCode == item.JobDurationCode) &&
                   (string.IsNullOrWhiteSpace(item.DurationHours) || this.DurationHours == item.DurationHours) &&
                   (string.IsNullOrWhiteSpace(item.StandardJob) || this.StandardJob == item.StandardJob) &&
                   (string.IsNullOrWhiteSpace(item.CorrectiveDescription) || this.CorrectiveDescription == item.CorrectiveDescription) &&
                   (string.IsNullOrWhiteSpace(item.WorkOrder) || this.WorkOrder == item.WorkOrder))
                    return true;
                //debugging
                /*
                var raisedTime = (string.IsNullOrWhiteSpace(item.RaisedTime) || this.RaisedTime == item.RaisedTime);
                var incidentDescription = (string.IsNullOrWhiteSpace(item.IncidentDescription) || this.IncidentDescription == item.IncidentDescription);
                var maintenanceType = (string.IsNullOrWhiteSpace(item.MaintenanceType) || this.MaintenanceType == item.MaintenanceType);
                var originator = (string.IsNullOrWhiteSpace(item.Originator) || this.Originator == item.Originator);
                var jobDurationFinish = (string.IsNullOrWhiteSpace(item.JobDurationFinish) || this.JobDurationFinish == item.JobDurationFinish);
                var incidentStatus = (string.IsNullOrWhiteSpace(item.IncidentStatus) || this.IncidentStatus == item.IncidentStatus);
                var equipmentReference = (string.IsNullOrWhiteSpace(item.EquipmentReference) || this.EquipmentReference == item.EquipmentReference);
                var compCode = (string.IsNullOrWhiteSpace(item.ComponentCode) || this.ComponentCode == item.ComponentCode);
                var modCode = (string.IsNullOrWhiteSpace(item.ModifierCode) || this.ModifierCode == item.ModifierCode);
                var jobDurationCode = (string.IsNullOrWhiteSpace(item.JobDurationCode) || this.JobDurationCode == item.JobDurationCode);
                var jobDurationHours = (string.IsNullOrWhiteSpace(item.DurationHours) || this.DurationHours == item.DurationHours);
                var standardJob = (string.IsNullOrWhiteSpace(item.StandardJob) || this.StandardJob == item.StandardJob);
                var correctiveDescription = (string.IsNullOrWhiteSpace(item.CorrectiveDescription) || this.CorrectiveDescription == item.CorrectiveDescription);
                var workOrder = (string.IsNullOrWhiteSpace(item.WorkOrder) || this.WorkOrder == item.WorkOrder);

                if (raisedTime &&
                    incidentDescription &&
                    maintenanceType &&
                    originator &&
                    jobDurationFinish &&
                    incidentStatus &&
                    equipmentReference &&
                    compCode &&
                    modCode &&
                    jobDurationCode &&
                    jobDurationHours &&
                    standardJob &&
                    correctiveDescription &&
                    workOrder)
                    return true;
                    */
            }
            else
            {
                if (this.RaisedTime == item.RaisedTime &&
                   this.IncidentDescription == item.IncidentDescription &&
                   this.MaintenanceType == item.MaintenanceType &&
                   this.Originator == item.Originator &&
                   this.JobDurationFinish == item.JobDurationFinish &&
                   this.IncidentStatus == item.IncidentStatus &&
                   this.EquipmentReference == item.EquipmentReference &&
                   this.ComponentCode == item.ComponentCode &&
                   this.ModifierCode == item.ModifierCode &&
                   this.JobDurationCode == item.JobDurationCode &&
                   this.DurationHours == item.DurationHours &&
                   this.StandardJob == item.StandardJob &&
                   this.CorrectiveDescription == item.CorrectiveDescription &&
                   this.WorkOrder == item.WorkOrder)
                    return true;
            }
            return false;
        }
    }
}
