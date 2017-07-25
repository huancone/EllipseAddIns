
using EllipseCommonsClassLibrary;

namespace EllipseWorkOrdersClassLibrary
{
    public class WorkOrderQualityStyles
    {
        public string WorkGroup;
        public string EquipmentNo;
        public string WorkOrderStatusM;//Estado de orden (si fue reabierta)
        public string WorkOrderStatusU;//User status
        public string CompCode;
        public string WorkOrderType;
        public string MaintenanceType;
        public string OriginatorPriority;
        public string PlanPriority;
        public string UnitOfWork;
        public string UnitsRequired;
        public string UnitsCompleted;
        public string ActualDurationHrs;
        public string ActualLabHrs;
        public string ActualLabCost;
        public string ActualMatCost;
        public string ActualOtherCost;
        public string JobCodesFlag;
        public string CompleteTextFlag;

        public WorkOrderQualityStyles(WorkOrder wo)
        {
            WorkGroup = string.IsNullOrWhiteSpace(wo.workGroup) ? StyleConstants.Warning : StyleConstants.Normal;
            EquipmentNo = !wo.maintenanceType.Equals("NM") && string.IsNullOrWhiteSpace(wo.equipmentNo) ? StyleConstants.Warning : StyleConstants.Normal;
            WorkOrderStatusM = wo.workOrderStatusM != "C" && !string.IsNullOrWhiteSpace(wo.completedCode) ? StyleConstants.Warning : StyleConstants.Normal;
            CompCode = wo.workOrderType.Equals("RE") && string.IsNullOrWhiteSpace(wo.compCode) ? StyleConstants.Error : StyleConstants.Normal;
            WorkOrderType = !WoTypeMtType.ValidateWoMtTypeCode(wo.workOrderType, wo.maintenanceType) ? StyleConstants.Error : StyleConstants.Normal;
            MaintenanceType = !WoTypeMtType.ValidateWoMtTypeCode(wo.workOrderType, wo.maintenanceType) ? StyleConstants.Error : StyleConstants.Normal;
            WorkOrderStatusU = wo.workOrderStatusM != "C" && string.IsNullOrEmpty(wo.workOrderStatusU) && !WorkOrderActions.ValidateUserStatus(wo.raisedDate, 60) ? StyleConstants.Warning : StyleConstants.Normal;
            OriginatorPriority = !WoTypeMtType.ValidatePriority(wo.origPriority) ? StyleConstants.Error : StyleConstants.Normal;
            PlanPriority = !WoTypeMtType.ValidatePriority(wo.origPriority) ? StyleConstants.Error : StyleConstants.Normal;
            int result;
            var unitsRequired = int.TryParse(wo.unitsRequired, out result) ? int.Parse(wo.unitsRequired) : 0;
            var unitsCompleted = int.TryParse(wo.unitsRequired, out result) ? int.Parse(wo.unitsComplete) : 0;
            UnitOfWork = !string.IsNullOrWhiteSpace(wo.unitOfWork) != (unitsRequired > 0) ? StyleConstants.Warning : StyleConstants.Normal;
            UnitsRequired = !string.IsNullOrWhiteSpace(wo.unitOfWork) != (unitsRequired > 0) ? StyleConstants.Warning : StyleConstants.Normal;
            UnitsCompleted = unitsRequired > 0 && unitsCompleted < unitsRequired ? StyleConstants.Warning : StyleConstants.Normal;

            var warningStyle = StyleConstants.Warning;
            if (!wo.workOrderType.Equals("RE") && !string.IsNullOrWhiteSpace(wo.stdJobNo) && !wo.maintenanceType.Equals("NM"))
                warningStyle = StyleConstants.Error;

            var estimateDurHrs = wo.estimatedDurationsHrs;
            var estimateLabHrs = (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabHrs : wo.estimatedLabHrs);
            var estimateLabCost = (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabCost : wo.estimatedLabCost);
            var estimateMatCost = (wo.calculatedMatFlag.Equals("Y") ? wo.calculatedMatCost : wo.estimatedMatCost);

            ActualDurationHrs = StyleConstants.Normal;
            ActualLabHrs = StyleConstants.Normal;
            ActualLabCost = StyleConstants.Normal;
            ActualMatCost = StyleConstants.Normal;
            ActualOtherCost = StyleConstants.Normal;
            //durationHrs
            if (!MathUtil.InThreshold(estimateDurHrs, wo.actualDurationsHrs, 1f))
                ActualDurationHrs = StyleConstants.Error;
            else if (!MathUtil.InThreshold(estimateDurHrs, wo.actualDurationsHrs, .2f))
                ActualDurationHrs = warningStyle;
            //lab hrs
            if (!MathUtil.InThreshold(estimateLabHrs, wo.actualLabHrs, 1f))
                ActualLabHrs = StyleConstants.Error;
            else if (!MathUtil.InThreshold(estimateLabHrs, wo.actualLabHrs, .2f))
                ActualLabHrs = warningStyle;
            //lab cost
            if (!MathUtil.InThreshold(estimateLabCost, wo.actualLabCost, 1f))
                ActualLabCost = StyleConstants.Error;
            else if (!MathUtil.InThreshold(estimateLabCost, wo.actualLabCost, .2f))
                ActualLabCost = warningStyle;
            //mat cost
            if (!MathUtil.InThreshold(estimateMatCost, wo.actualMatCost, 1f))
                ActualMatCost = StyleConstants.Error;
            else if (!MathUtil.InThreshold(estimateMatCost, wo.actualMatCost, .2f))
                ActualMatCost = warningStyle;
            //other cost
            if (!MathUtil.InThreshold(wo.estimatedOtherCost, wo.actualOtherCost, 1f))
                ActualOtherCost = StyleConstants.Error;
            else if (!MathUtil.InThreshold(wo.estimatedOtherCost, wo.actualOtherCost, .2f))
                ActualOtherCost = warningStyle;


            JobCodesFlag = wo.maintenanceType.Equals("CO") && !wo.jobCodeFlag.Equals("Y") ? StyleConstants.Error : StyleConstants.Normal;
            CompleteTextFlag = wo.completeTextFlag == "N" ? StyleConstants.Warning : StyleConstants.Normal;
        }

    }
}
