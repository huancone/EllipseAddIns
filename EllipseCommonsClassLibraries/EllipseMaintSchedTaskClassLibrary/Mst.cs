using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using EllipseCommonsClassLibrary.PrimitiveClasses;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseMaintSchedTaskClassLibrary
{
    [XmlRoot(ElementName = "dto")]
    public class Mst : TaskMwpDto
    {
        [XmlElement("addedToOutage")] public string AddedToOutage;

        [XmlElement("autoReqInd")] public string AutoReqInd;

        [XmlElement("commitStartDate")] public string CommitStartDate;

        [XmlElement("commitStartTime")] public string CommitStartTime;

        [XmlElement("fixedScheduling")] public string FixedScheduling;

        [XmlElement("forceConversion")] public string ForceConversion;

        [XmlElement("overriding")] public string Overriding;

        [XmlElement("plannedStartDate")] public string PlannedStartDate;

        [XmlElement("plannedStartTime")] public string PlannedStartTime;

        [XmlElement("printJobCard")] public string PrintJobCard;

        [XmlElement("schedInd700")] public string SchedInd700;

        [XmlElement("toleranceDays")] public string ToleranceDays;

        [XmlElement("tolerancePc")] public string TolerancePc;

        [XmlElement("unitsScale")] public string UnitsScale;

        public Mst()
        {

        }
        public Mst(MstService.MSTiMWPDTO mstiMwpDto)
        {

             AddedToOutage = mstiMwpDto.addedToOutageSpecified ? MyUtilities.ToString(mstiMwpDto.addedToOutage) : null;

             AutoReqInd = mstiMwpDto.autoReqIndSpecified ? MyUtilities.ToString(mstiMwpDto.autoReqInd) : null;

             CommitStartDate = mstiMwpDto.commitStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.commitStartDate) : null;

             CommitStartTime = mstiMwpDto.commitStartDateSpecified ? mstiMwpDto.commitStartTime : null;

             FixedScheduling = mstiMwpDto.fixedScheduling;

             ForceConversion = mstiMwpDto.forceConversionSpecified ? MyUtilities.ToString(mstiMwpDto.forceConversion) : null;

             Overriding = mstiMwpDto.overridingSpecified ? MyUtilities.ToString(mstiMwpDto.overriding) : null;

             PlannedStartDate = mstiMwpDto.plannedStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.plannedStartDate) : null;

             PlannedStartTime = mstiMwpDto.plannedStartDateSpecified ? mstiMwpDto.plannedStartTime : null;

             PrintJobCard = mstiMwpDto.printJobCardSpecified ? MyUtilities.ToString(mstiMwpDto.printJobCard) : null;
                 
             SchedInd700 = mstiMwpDto.schedInd700;

             ToleranceDays = mstiMwpDto.toleranceDaysSpecified ? MyUtilities.ToString(mstiMwpDto.toleranceDays) : null;

             TolerancePc = mstiMwpDto.tolerancePCSpecified ? MyUtilities.ToString(mstiMwpDto.tolerancePC) : null;
                 
             UnitsScale = mstiMwpDto.unitsScaleSpecified ? MyUtilities.ToString(mstiMwpDto.unitsScale) : null;


        //Herencia de Task
            AplCompCode = mstiMwpDto.APLCompCode;

            AplEgi = mstiMwpDto.APLEGI;

            AplEquipment = mstiMwpDto.APLEquipment;

            AplModCode = mstiMwpDto.APLModCode;

            AplSequenceNumber = mstiMwpDto.APLSequenceNumber;

            AplType = mstiMwpDto.APLType;

            AssignPersonForTasks = mstiMwpDto.assignPersonForTasks;

            AssignPersonName = mstiMwpDto.assignPersonName;

            CalculatedTotalCosts = mstiMwpDto.calculatedTotalCostsSpecified ? MyUtilities.ToString(mstiMwpDto.calculatedTotalCosts) : null;

            CompletionComments = mstiMwpDto.completionComments;

            CompletionCommentsHeader = mstiMwpDto.completionCommentsHeader;

            CompletionInstruction = mstiMwpDto.completionInstruction;

            CrewForTasks = mstiMwpDto.crewForTasks;

            CrewTypeForTasks = mstiMwpDto.crewTypeForTasks;

            EarliestFinDate = mstiMwpDto.earliestFinDateSpecified ? MyUtilities.ToString(mstiMwpDto.earliestFinDate) : null;

            EarliestFinTime = mstiMwpDto.earliestFinTime;

            EarliestStrDate = mstiMwpDto.earliestStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.earliestStrDate) : null;

            EarliestStrTime = mstiMwpDto.earliestStrTime;

            EstimatedLabourHours = mstiMwpDto.estimatedLabourHoursSpecified ? MyUtilities.ToString(mstiMwpDto.estimatedLabourHours) : null;

            EstimatedTotalCosts = mstiMwpDto.estimatedTotalCostsSpecified ? MyUtilities.ToString(mstiMwpDto.estimatedTotalCosts) : null;

            FloatDays = mstiMwpDto.floatDays;

            JobDescCode = mstiMwpDto.jobDescCode;

            JobDescCodeDescription = mstiMwpDto.jobDescCodeDescription;

            JobInstructions = mstiMwpDto.jobInstructions;

            LatestStrDate = mstiMwpDto.latestStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.latestStrDate) : null;

            LatestStrTime = mstiMwpDto.latestStrTime;

            LinkedTaskInd = mstiMwpDto.linkedTaskInd;

            MachineHoursEstimated = mstiMwpDto.machineHoursEstimatedSpecified ? MyUtilities.ToString(mstiMwpDto.machineHoursEstimated) : null;

            ResourceGroups = mstiMwpDto.resourceGroups;

            RestartChildTaskNumber = mstiMwpDto.restartChildTaskNumber;

            RestartParentTaskNumber = mstiMwpDto.restartParentTaskNumber;

            SafetyInstruction = mstiMwpDto.safetyInstruction;

            ScheduledUnitsDay = mstiMwpDto.scheduledUnitsDaySpecified ? MyUtilities.ToString(mstiMwpDto.scheduledUnitsDay) : null;

            StatusDescription = mstiMwpDto.statusDescription;

            TaskActualFinishDate = mstiMwpDto.taskActualFinishDateSpecified ? MyUtilities.ToString(mstiMwpDto.taskActualFinishDate) : null;

            TaskActualFinishTime = mstiMwpDto.taskActualFinishTime;

            TaskActualStartDate = mstiMwpDto.taskActualStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.taskActualStartDate) : null;

            TaskActualStartTime = mstiMwpDto.taskActualStartTime;

            TaskComplete = mstiMwpDto.taskCompleteSpecified ? MyUtilities.ToString(mstiMwpDto.taskComplete) : null;

            TaskCompletionText = mstiMwpDto.taskCompletionText;

            TaskCreationDate = mstiMwpDto.taskCreationDateSpecified ? MyUtilities.ToString(mstiMwpDto.taskCreationDate) : null;

            TaskCreationTime = mstiMwpDto.taskCreationTime;

            TaskDescription = mstiMwpDto.taskDescription;

            TaskEffort = mstiMwpDto.taskEffort;

            TaskJobDescription = mstiMwpDto.taskJobDescription;

            TaskLocationFrom = mstiMwpDto.taskLocationFrom;

            TaskLocationTo = mstiMwpDto.taskLocationTo;

            TaskNotes = mstiMwpDto.taskNotes;

            TaskPlannerPriority = mstiMwpDto.taskPlannerPriority;

            TaskResources = mstiMwpDto.taskResources;

            TaskUnitsComplete = mstiMwpDto.taskUnitsCompleteSpecified ? MyUtilities.ToString(mstiMwpDto.taskUnitsComplete) : null;

            TaskUnitsRequired = mstiMwpDto.taskUnitsRequiredSpecified ? MyUtilities.ToString(mstiMwpDto.taskUnitsRequired) : null;

            TemplateType = mstiMwpDto.templateType;

            UncompletedTodos = mstiMwpDto.uncompletedTodosSpecified ? MyUtilities.ToString(mstiMwpDto.uncompletedTodos) : null;

            UnitsPerDay = mstiMwpDto.unitsPerDaySpecified ? MyUtilities.ToString(mstiMwpDto.unitsPerDay) : null;

            UserStatus = mstiMwpDto.userStatus;

            UserStatusDescription = mstiMwpDto.userStatusDescription;

            WoTaskDesc = mstiMwpDto.woTaskDesc;

            WorkActivityClassification = mstiMwpDto.workActivityClassification;

            WorkCentre = mstiMwpDto.workCentre;

            WorkGroupForTasks = mstiMwpDto.workGroupForTasks;

            WorkRelease = mstiMwpDto.workReleaseSpecified ? MyUtilities.ToString(mstiMwpDto.workRelease) : null;

            WorkReleasePrior = mstiMwpDto.workReleasePriorSpecified ? MyUtilities.ToString(mstiMwpDto.workReleasePrior) : null;

            WoTaskNo = mstiMwpDto.WOTaskNo;

            AccountCode = mstiMwpDto.accountCode;

            AccountCodeDescription = mstiMwpDto.accountCodeDescription;

            ActDurHrs = mstiMwpDto.actDurHrsSpecified ? MyUtilities.ToString(mstiMwpDto.actDurHrs) : null;

            ActEquipCost = mstiMwpDto.actEquipCostSpecified ? MyUtilities.ToString(mstiMwpDto.actEquipCost) : null;

            ActLabCost = mstiMwpDto.actLabCostSpecified ? MyUtilities.ToString(mstiMwpDto.actLabCost) : null;

            ActLabHrs = mstiMwpDto.actLabHrsSpecified ? MyUtilities.ToString(mstiMwpDto.actLabHrs) : null;

            ActMatCost = mstiMwpDto.actMatCostSpecified ? MyUtilities.ToString(mstiMwpDto.actMatCost) : null;

            ActOtherCost = mstiMwpDto.actOtherCostSpecified ? MyUtilities.ToString(mstiMwpDto.actOtherCost) : null;

            ActualCostReallocation = mstiMwpDto.actualCostReallocationSpecified ? MyUtilities.ToString(mstiMwpDto.actualCostReallocation) : null;

            ActualFinishDate = mstiMwpDto.actualFinishDateSpecified ? MyUtilities.ToString(mstiMwpDto.actualFinishDate) : null;

            ActualFinishTime = mstiMwpDto.actualFinishTime;

            ActualStartDate = mstiMwpDto.actualStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.actualStartDate) : null;

            ActualStartTime = mstiMwpDto.actualStartTime;

            ActualTotalCost = mstiMwpDto.actualTotalCostSpecified ? MyUtilities.ToString(mstiMwpDto.actualTotalCost) : null;

            AptwExistsSw = mstiMwpDto.aptwExistsSwSpecified ? MyUtilities.ToString(mstiMwpDto.aptwExistsSw) : null;

            AssignPerson = mstiMwpDto.assignPerson;

            AssocEquipSw = mstiMwpDto.assocEquipSwSpecified ? MyUtilities.ToString(mstiMwpDto.assocEquipSw) : null;

            AssocEquipmentItemNo = mstiMwpDto.assocEquipmentItemNo;

            AssociatedEquipment = mstiMwpDto.associatedEquipment;

            AssumeFirstMSTI = mstiMwpDto.assumeFirstMSTISpecified ? MyUtilities.ToString(mstiMwpDto.assumeFirstMSTI) : null;

            AuthsdBy = mstiMwpDto.authsdBy;

            BillableInd = mstiMwpDto.billableIndSpecified ? MyUtilities.ToString(mstiMwpDto.billableInd) : null;

            BillingLvlInd = mstiMwpDto.billingLvlInd;

            CalcEquipCost = mstiMwpDto.calcEquipCostSpecified ? MyUtilities.ToString(mstiMwpDto.calcEquipCost) : null;

            CalcLabCost = mstiMwpDto.calcLabCostSpecified ? MyUtilities.ToString(mstiMwpDto.calcLabCost) : null;

            CalcLabHrs = mstiMwpDto.calcLabHrsSpecified ? MyUtilities.ToString(mstiMwpDto.calcLabHrs) : null;

            CalcMatCost = mstiMwpDto.calcMatCostSpecified ? MyUtilities.ToString(mstiMwpDto.calcMatCost) : null;

            CalcOthCost = mstiMwpDto.calcOthCostSpecified ? MyUtilities.ToString(mstiMwpDto.calcOthCost) : null;

            CalculatedEquipmentFlag = mstiMwpDto.calculatedEquipmentFlagSpecified ? MyUtilities.ToString(mstiMwpDto.calculatedEquipmentFlag) : null;

            CalculatedLabFlag = mstiMwpDto.calculatedLabFlagSpecified ? MyUtilities.ToString(mstiMwpDto.calculatedLabFlag) : null;

            CalculatedMatFlag = mstiMwpDto.calculatedMatFlagSpecified ? MyUtilities.ToString(mstiMwpDto.calculatedMatFlag) : null;

            CalculatedOtherFlag = mstiMwpDto.calculatedOtherFlagSpecified ? MyUtilities.ToString(mstiMwpDto.calculatedOtherFlag) : null;

            CalculatedTotalFlag = mstiMwpDto.calculatedTotalFlagSpecified ? MyUtilities.ToString(mstiMwpDto.calculatedTotalFlag) : null;

            ClosedDt = mstiMwpDto.closedDtSpecified ? MyUtilities.ToString(mstiMwpDto.closedDt) : null;

            ClosedStatus = mstiMwpDto.closedStatus;

            ClosedTime = mstiMwpDto.closedTimeSpecified ? MyUtilities.ToString(mstiMwpDto.closedTime) : null;

            CompCode = mstiMwpDto.compCode;

            CompModCode = mstiMwpDto.compModCode;

            CompletedBy = mstiMwpDto.completedBy;

            CompletedCode = mstiMwpDto.completedCode;

            CompletionText = mstiMwpDto.completionText;

            CompletionTextExists = mstiMwpDto.completionTextExistsSpecified ? MyUtilities.ToString(mstiMwpDto.completionTextExists) : null;

            ConAstSegFr = mstiMwpDto.conAstSegFrSpecified ? MyUtilities.ToString(mstiMwpDto.conAstSegFr) : null;

            ConAstSegLength = mstiMwpDto.conAstSegLengthSpecified ? MyUtilities.ToString(mstiMwpDto.conAstSegLength) : null;

            ConAstSegTo = mstiMwpDto.conAstSegToSpecified ? MyUtilities.ToString(mstiMwpDto.conAstSegTo) : null;

            CountyShire = mstiMwpDto.countyShire;

            CreationDate = mstiMwpDto.creationDateSpecified ? MyUtilities.ToString(mstiMwpDto.creationDate) : null;

            CreationTime = mstiMwpDto.creationDateSpecified ? mstiMwpDto.creationTime : null;

            Crew = mstiMwpDto.crew;

            CrteInsitu = mstiMwpDto.crteInsituSpecified ? MyUtilities.ToString(mstiMwpDto.crteInsitu) : null;

            CurrentStatDate1 = mstiMwpDto.currentStatDate1Specified ? MyUtilities.ToString(mstiMwpDto.currentStatDate1) : null;

            CurrentStatDate2 = mstiMwpDto.currentStatDate2Specified ? MyUtilities.ToString(mstiMwpDto.currentStatDate2) : null;

            CurrentStatType1 = mstiMwpDto.currentStatType1;

            CurrentStatType2 = mstiMwpDto.currentStatType2;

            CurrentStatValue1 = mstiMwpDto.currentStatValue1Specified ? MyUtilities.ToString(mstiMwpDto.currentStatValue1) : null;

            CurrentStatValue2 = mstiMwpDto.currentStatValue2Specified ? MyUtilities.ToString(mstiMwpDto.currentStatValue2) : null;



            Data1732 = mstiMwpDto.data1732;

            DateStatus = mstiMwpDto.dateStatus;

            DstrctAcctCode = mstiMwpDto.dstrctAcctCode;

            DstrctCode = mstiMwpDto.dstrctCode;

            EmailAddress = mstiMwpDto.emailAddress;

            EquipClass = mstiMwpDto.equipClass;

            EquipClassifx1 = mstiMwpDto.equipClassifx1;

            EquipClassifx10 = mstiMwpDto.equipClassifx10;

            EquipClassifx11 = mstiMwpDto.equipClassifx11;

            EquipClassifx12 = mstiMwpDto.equipClassifx12;

            EquipClassifx13 = mstiMwpDto.equipClassifx13;

            EquipClassifx14 = mstiMwpDto.equipClassifx14;

            EquipClassifx15 = mstiMwpDto.equipClassifx15;

            EquipClassifx16 = mstiMwpDto.equipClassifx16;

            EquipClassifx17 = mstiMwpDto.equipClassifx17;

            EquipClassifx18 = mstiMwpDto.equipClassifx18;

            EquipClassifx19 = mstiMwpDto.equipClassifx19;

            EquipClassifx2 = mstiMwpDto.equipClassifx2;

            EquipClassifx20 = mstiMwpDto.equipClassifx20;

            EquipClassifx3 = mstiMwpDto.equipClassifx3;

            EquipClassifx4 = mstiMwpDto.equipClassifx4;

            EquipClassifx5 = mstiMwpDto.equipClassifx5;

            EquipClassifx6 = mstiMwpDto.equipClassifx6;

            EquipClassifx7 = mstiMwpDto.equipClassifx7;

            EquipClassifx8 = mstiMwpDto.equipClassifx8;

            EquipClassifx9 = mstiMwpDto.equipClassifx9;

            EquipGrpId = mstiMwpDto.equipGrpId;

            EquipLocation = mstiMwpDto.equipLocation;

            EquipNo = mstiMwpDto.equipNo;

            EquipStatus = mstiMwpDto.equipStatus;

            EquipUpdateFlag = mstiMwpDto.equipUpdateFlagSpecified ? MyUtilities.ToString(mstiMwpDto.equipUpdateFlag) : null;

            EquipmentClassDescription = mstiMwpDto.equipmentClassDescription;

            EstDurHrs = mstiMwpDto.estDurHrsSpecified ? MyUtilities.ToString(mstiMwpDto.estDurHrs) : null;

            EstEquipCost = mstiMwpDto.estEquipCostSpecified ? MyUtilities.ToString(mstiMwpDto.estEquipCost) : null;

            EstLabCost = mstiMwpDto.estLabCostSpecified ? MyUtilities.ToString(mstiMwpDto.estLabCost) : null;

            EstLabHrs = mstiMwpDto.estLabHrsSpecified ? MyUtilities.ToString(mstiMwpDto.estLabHrs) : null;

            EstMatCost = mstiMwpDto.estMatCostSpecified ? MyUtilities.ToString(mstiMwpDto.estMatCost) : null;

            EstOtherCost = mstiMwpDto.estOtherCostSpecified ? MyUtilities.ToString(mstiMwpDto.estOtherCost) : null; ;

            EstimateDescription = mstiMwpDto.estimateDescription;

            EstimateNo = mstiMwpDto.estimateNo;

            EstimatedTotalCost = mstiMwpDto.estimatedTotalCostSpecified ? MyUtilities.ToString(mstiMwpDto.estimatedTotalCost) : null;

            ExistingPlannedFinishDate = mstiMwpDto.existingPlannedFinishDateSpecified ? MyUtilities.ToString(mstiMwpDto.existingPlannedFinishDate) : null;

            ExistingPlannedFinishTime = mstiMwpDto.existingPlannedFinishTime;

            ExistingPlannedStartDate = mstiMwpDto.existingPlannedStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.existingPlannedStartDate) : null;

            ExistingPlannedStartTime = mstiMwpDto.existingPlannedStartTime;

            ExtendedText = mstiMwpDto.extendedText;

            ExtendedTextExists = mstiMwpDto.extendedTextExistsSpecified ? MyUtilities.ToString(mstiMwpDto.extendedTextExists) : null;

            FailurePart = mstiMwpDto.failurePart;

            FaxNumber = mstiMwpDto.faxNumber;

            FinalCostIndicator = mstiMwpDto.finalCostIndicatorSpecified ? MyUtilities.ToString(mstiMwpDto.finalCostIndicator) : null;

            FromLink = mstiMwpDto.fromLinkSpecified ? MyUtilities.ToString(mstiMwpDto.fromLink) : null;

            GanttFinishDateTime = mstiMwpDto.ganttFinishDateTime;

            GanttLinkId = mstiMwpDto.ganttLinkId;

            GanttParentLinkId = mstiMwpDto.ganttParentLinkId;

            GanttStartDateTime = mstiMwpDto.ganttStartDateTime;

            HasAssignTodo = mstiMwpDto.hasAssignTodo;

            Id = mstiMwpDto.idSpecified ? MyUtilities.ToString(mstiMwpDto.id) : null;

            ImmediateInspections = mstiMwpDto.immediateInspectionsSpecified ? MyUtilities.ToString(mstiMwpDto.immediateInspections) : null;

            IsControl = mstiMwpDto.isControlSpecified ? MyUtilities.ToString(mstiMwpDto.isControl) : null;

            IsMSTParent = mstiMwpDto.isMSTParent;

            IsParent = mstiMwpDto.isParent;

            ItemName1 = mstiMwpDto.itemName1;

            ItemName2 = mstiMwpDto.itemName2;

            JobId = mstiMwpDto.jobId;

            JobParentId = mstiMwpDto.jobParentId;

            JobType = mstiMwpDto.jobType;

            LastModifiedDate = mstiMwpDto.lastModifiedDateSpecified ? MyUtilities.ToString(mstiMwpDto.lastModifiedDate) : null;

            LastModifiedTime = mstiMwpDto.lastModifiedTime;

            LastPerformedDate = mstiMwpDto.lastPerformedDateSpecified ? MyUtilities.ToString(mstiMwpDto.lastPerformedDate) : null;

            LastPerformedDstrctCode = mstiMwpDto.lastPerformedDstrctCode;

            LastPerformedWorkOrder = mstiMwpDto.lastPerformedWorkOrder;

            LastTranRloc = mstiMwpDto.lastTranRloc;

            LinkDistrictCode = mstiMwpDto.linkDistrictCode;

            LinkEntityId = mstiMwpDto.linkEntityIdSpecified ? MyUtilities.ToString(mstiMwpDto.linkEntityId) : null;

            LinkEntityType = mstiMwpDto.linkEntityTypeSpecified ? MyUtilities.ToString(mstiMwpDto.linkEntityType) : null;

            LinkErpRef = mstiMwpDto.linkErpRef;

            LinkErpTaskRef = mstiMwpDto.linkErpTaskRef;

            LinkId = mstiMwpDto.linkIdSpecified ? MyUtilities.ToString(mstiMwpDto.linkId) : null;

            LinkJobLinkListId = mstiMwpDto.linkJobLinkListIdSpecified ? MyUtilities.ToString(mstiMwpDto.linkJobLinkListId) : null;

            LinkLagOrLead = mstiMwpDto.linkLagOrLead;

            LinkOffset = mstiMwpDto.linkOffsetSpecified ? MyUtilities.ToString(mstiMwpDto.linkOffset) : null;

            LinkOffsetType = mstiMwpDto.linkOffsetType;

            LinkScale = mstiMwpDto.linkScale;

            LinkScaledOffset = mstiMwpDto.linkScaledOffsetSpecified ? MyUtilities.ToString(mstiMwpDto.linkScaledOffset) : null;

            LinkScheduleType = mstiMwpDto.linkScheduleType;

            LinkSequenceNum = mstiMwpDto.linkSequenceNumSpecified ? MyUtilities.ToString(mstiMwpDto.linkSequenceNum) : null;

            LinkStatType = mstiMwpDto.linkStatType;

            LinkStatValue = mstiMwpDto.linkStatValueSpecified ? MyUtilities.ToString(mstiMwpDto.linkStatValue) : null;

            LinkState = mstiMwpDto.linkState;

            LinkType = mstiMwpDto.linkType;

            LinkedInd = mstiMwpDto.linkedInd;

            Location = mstiMwpDto.location;

            LocationFr = mstiMwpDto.locationFr;

            MaintSchTask = mstiMwpDto.maintSchTask;

            MaintType = mstiMwpDto.maintType;

            MaintTypeDescription = mstiMwpDto.maintTypeDescription;

            MatUpdateFlag = mstiMwpDto.matUpdateFlagSpecified ? MyUtilities.ToString(mstiMwpDto.matUpdateFlag) : null;

            Mnemonic = mstiMwpDto.mnemonic;

            MstReference = mstiMwpDto.mstReference;

            MultipleResourceRequirements = mstiMwpDto.multipleResourceRequirements;

            MustStartInd = mstiMwpDto.mustStartInd;

            NoGanttMove = mstiMwpDto.noGanttMove;

            NoGanttResize = mstiMwpDto.noGanttResize;

            NoOfTasks = mstiMwpDto.noOfTasks;

            NoTasksCompl = mstiMwpDto.noTasksCompl;

            NumberOfMSTisRemaining = mstiMwpDto.numberOfMSTisRemainingSpecified ? MyUtilities.ToString(mstiMwpDto.numberOfMSTisRemaining) : null;

            OrigPriority = mstiMwpDto.origPriority;

            OrigSchedDate = mstiMwpDto.origSchedDateSpecified ? MyUtilities.ToString(mstiMwpDto.origSchedDate) : null;

            OriginalPlannedFinishDate = mstiMwpDto.originalPlannedFinishDateSpecified ? MyUtilities.ToString(mstiMwpDto.originalPlannedFinishDate) : null;

            OriginalPlannedFinishTime = mstiMwpDto.originalPlannedFinishTime;

            OriginalPlannedStartDate = mstiMwpDto.originalPlannedStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.originalPlannedStartDate) : null;

            OriginalPlannedStartTime = mstiMwpDto.originalPlannedStartTime;

            OriginatorId = mstiMwpDto.originatorId;

            OtherUpdateFlag = mstiMwpDto.otherUpdateFlagSpecified ? MyUtilities.ToString(mstiMwpDto.otherUpdateFlag) : null;

            OutServDate = mstiMwpDto.outServDateSpecified ? MyUtilities.ToString(mstiMwpDto.outServDate) : null;

            OutServTime = mstiMwpDto.outServTime;

            OutageDescription = mstiMwpDto.outageDescription;

            OutageReference = mstiMwpDto.outageReference;

            OutageStatus = mstiMwpDto.outageStatus;

            PaperHist = mstiMwpDto.paperHist;

            ParentEntityId = mstiMwpDto.parentEntityIdSpecified ? MyUtilities.ToString(mstiMwpDto.parentEntityId) : null;

            ParentEquip = mstiMwpDto.parentEquip;

            ParentJobType = mstiMwpDto.parentJobType;

            ParentLinkId = mstiMwpDto.parentLinkIdSpecified ? MyUtilities.ToString(mstiMwpDto.parentLinkId) : null;

            ParentPlantNo = mstiMwpDto.parentPlantNo;

            ParentWo = mstiMwpDto.parentWo;

            ParentWoDescription = mstiMwpDto.parentWoDescription;

            PartNo = mstiMwpDto.partNo;

            PartialCacheKey = mstiMwpDto.partialCacheKey;

            PartsAvailabe = mstiMwpDto.partsAvailabe;

            PcComplete = mstiMwpDto.pcCompleteSpecified ? MyUtilities.ToString(mstiMwpDto.pcComplete) : null;

            PermitReqdSw = mstiMwpDto.permitReqdSwSpecified ? MyUtilities.ToString(mstiMwpDto.permitReqdSw) : null;

            PlanFinDate = mstiMwpDto.planFinDateSpecified ? MyUtilities.ToString(mstiMwpDto.planFinDate) : null;

            PlanFinTime = mstiMwpDto.planFinTime;

            PlanOffsetSw = mstiMwpDto.planOffsetSwSpecified ? MyUtilities.ToString(mstiMwpDto.planOffsetSw) : null;

            PlanPriority = mstiMwpDto.planPriority;

            PlanStatType = mstiMwpDto.planStatType;

            PlanStatVal = mstiMwpDto.planStatValSpecified ? MyUtilities.ToString(mstiMwpDto.planStatVal) : null;

            PlanStrDate = mstiMwpDto.planStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.planStrDate) : null;

            PlanStrTime = mstiMwpDto.planStrTime;

            PlantNo = mstiMwpDto.plantNo;

            PrefDate = mstiMwpDto.prefDateSpecified ? MyUtilities.ToString(mstiMwpDto.prefDate) : null;

            PrefTime = mstiMwpDto.prefTime;

            PrinterName = mstiMwpDto.printerName;

            ProdUnitItem = mstiMwpDto.prodUnitItemSpecified ? MyUtilities.ToString(mstiMwpDto.prodUnitItem) : null;

            ProjDesc = mstiMwpDto.projDesc;

            ProjectNo = mstiMwpDto.projectNo;

            QuoteValue = mstiMwpDto.quoteValueSpecified ? MyUtilities.ToString(mstiMwpDto.quoteValue) : null;

            RaisedDate = mstiMwpDto.raisedDateSpecified ? MyUtilities.ToString(mstiMwpDto.raisedDate) : null;

            RaisedTime = mstiMwpDto.raisedTimeSpecified ? MyUtilities.ToString(mstiMwpDto.raisedTime) : null;

            ReallocationCostAccount = mstiMwpDto.reallocationCostAccount;

            ReallocationCrEe = mstiMwpDto.reallocationCrEe;

            ReallocationFreqInd = mstiMwpDto.reallocationFreqInd;

            ReallocationLimitVal = mstiMwpDto.reallocationLimitValSpecified ? MyUtilities.ToString(mstiMwpDto.reallocationLimitVal) : null;

            ReallocationMarginPc = mstiMwpDto.reallocationMarginPcSpecified ? MyUtilities.ToString(mstiMwpDto.reallocationMarginPc) : null;

            ReallocationMethod = mstiMwpDto.reallocationMethod;

            ReallocationProject = mstiMwpDto.reallocationProject;

            ReallocationVarAccount = mstiMwpDto.reallocationVarAccount;

            ReallocationWo = mstiMwpDto.reallocationWo;

            RecallTimeHrs = mstiMwpDto.recallTimeHrsSpecified ? MyUtilities.ToString(mstiMwpDto.recallTimeHrs) : null;

            Reference = mstiMwpDto.reference;

            RelatedWo = mstiMwpDto.relatedWo;

            ReqByDate = mstiMwpDto.reqByDateSpecified ? MyUtilities.ToString(mstiMwpDto.reqByDate) : null;

            ReqByTime = mstiMwpDto.reqByTime;

            ReqStartDate = mstiMwpDto.reqStartDateSpecified ? MyUtilities.ToString(mstiMwpDto.reqStartDate) : null;

            ReqStartTime = mstiMwpDto.reqStartTime;

            RequestId = mstiMwpDto.requestId;

            ResUpdateFlag = mstiMwpDto.resUpdateFlagSpecified ? MyUtilities.ToString(mstiMwpDto.resUpdateFlag) : null;

            RespondedDate = mstiMwpDto.respondedDateSpecified ? MyUtilities.ToString(mstiMwpDto.respondedDate) : null;

            RespondedTime = mstiMwpDto.respondedTime;

            RestartChildWODstrctCode = mstiMwpDto.restartChildWODstrctCode;

            RestartChildWOPlanStrDate = mstiMwpDto.restartChildWOPlanStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.restartChildWOPlanStrDate) : null;


            RestartChildWOWorkOrder = mstiMwpDto.restartChildWOWorkOrder;

            RestartMSTIFromLinkMstReference = mstiMwpDto.restartMSTIFromLinkMstReference;

            RestartMSTIFromLinkPlanStrDate = mstiMwpDto.restartMSTIFromLinkPlanStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.restartMSTIFromLinkPlanStrDate) : null;

            RestartMSTIMstReference = mstiMwpDto.restartMSTIMstReference;

            RestartMSTIPlanStrDate = mstiMwpDto.restartMSTIPlanStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.restartMSTIPlanStrDate) : null;

            RestartMstiTaskNo = mstiMwpDto.restartMSTiTaskNo;

            RestartParentWODstrctCode = mstiMwpDto.restartParentWODstrctCode;

            RestartParentWOPlanStrDate = mstiMwpDto.restartParentWOPlanStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.restartParentWOPlanStrDate) : null;

            RestartParentWOWorkOrder = mstiMwpDto.restartParentWOWorkOrder;

            RevenueCode = mstiMwpDto.revenueCode;

            SchSegFr = mstiMwpDto.schSegFrSpecified ? MyUtilities.ToString(mstiMwpDto.schSegFr) : null;

            SchSegLength = mstiMwpDto.schSegLengthSpecified ? MyUtilities.ToString(mstiMwpDto.schSegLength) : null;

            SchSegTo = mstiMwpDto.schSegToSpecified ? MyUtilities.ToString(mstiMwpDto.schSegTo) : null;

            SchedDesc1 = mstiMwpDto.schedDesc1;

            SchedDesc2 = mstiMwpDto.schedDesc2;

            SegmentUom = mstiMwpDto.segmentUom;

            SequenceNo = mstiMwpDto.sequenceNoSpecified ? MyUtilities.ToString(mstiMwpDto.sequenceNo) : null;

            SerialNumber = mstiMwpDto.serialNumber;

            ServiceOffDate = mstiMwpDto.serviceOffDateSpecified ? MyUtilities.ToString(mstiMwpDto.serviceOffDate) : null;

            ServiceOffTime = mstiMwpDto.serviceOffTime;

            ServiceOnDate = mstiMwpDto.serviceOnDateSpecified ? MyUtilities.ToString(mstiMwpDto.serviceOnDate) : null;

            ServiceOnTime = mstiMwpDto.serviceOnTime;

            ShortDesc1 = mstiMwpDto.shortDesc1;

            ShortDesc2 = mstiMwpDto.shortDesc2;

            ShutdownEquipmentNo = mstiMwpDto.shutdownEquipmentNo;

            ShutdownEquipmentRef = mstiMwpDto.shutdownEquipmentRef;

            ShutdownNo = mstiMwpDto.shutdownNo;

            ShutdownType = mstiMwpDto.shutdownType;

            Source = mstiMwpDto.source;

            State = mstiMwpDto.state;

            StatutoryFlg = mstiMwpDto.statutoryFlg;

            StdJobNo = mstiMwpDto.stdJobNo;

            StreetName = mstiMwpDto.streetName;

            StreetNo = mstiMwpDto.streetNo;

            Suburb = mstiMwpDto.suburb;

            SuppressingMST = mstiMwpDto.suppressingMST;

            TargtFinDate = mstiMwpDto.targtFinDateSpecified ? MyUtilities.ToString(mstiMwpDto.targtFinDate) : null;

            TargtStrDate = mstiMwpDto.targtStrDateSpecified ? MyUtilities.ToString(mstiMwpDto.targtStrDate) : null;

            TaskAptwSw = mstiMwpDto.taskAptwSw;

            TownCity = mstiMwpDto.townCity;

            UnitOfWork = mstiMwpDto.unitOfWork;

            UnitsComplete = mstiMwpDto.unitsCompleteSpecified ? MyUtilities.ToString(mstiMwpDto.unitsComplete) : null;

            UnitsRequired = mstiMwpDto.unitsRequiredSpecified ? MyUtilities.ToString(mstiMwpDto.unitsRequired) : null;

            UpperCostLimit = mstiMwpDto.upperCostLimitSpecified ? MyUtilities.ToString(mstiMwpDto.upperCostLimit) : null;

            WoDesc = mstiMwpDto.woDesc;

            WoJobCodex1 = mstiMwpDto.woJobCodex1;

            WoJobCodex10 = mstiMwpDto.woJobCodex10;

            WoJobCodex2 = mstiMwpDto.woJobCodex2;

            WoJobCodex3 = mstiMwpDto.woJobCodex3;

            WoJobCodex4 = mstiMwpDto.woJobCodex4;

            WoJobCodex5 = mstiMwpDto.woJobCodex5;

            WoJobCodex6 = mstiMwpDto.woJobCodex6;

            WoJobCodex7 = mstiMwpDto.woJobCodex7;

            WoJobCodex8 = mstiMwpDto.woJobCodex8;

            WoJobCodex9 = mstiMwpDto.woJobCodex9;

            WoStatusM = mstiMwpDto.woStatusM;

            WoStatusMDescription = mstiMwpDto.woStatusMDescription;

            WoStatusU = mstiMwpDto.woStatusU;

            WoStatusUDescription = mstiMwpDto.woStatusUDescription;

            WoType = mstiMwpDto.woType;

            WoTypeDescription = mstiMwpDto.woTypeDescription;

            WorkGroup = mstiMwpDto.workGroup;

            WorkOrder = mstiMwpDto.workOrder;

            WorkRequestDescription = mstiMwpDto.workRequestDescription;

            WorkRequestNumber = mstiMwpDto.workRequestNumber;

            ZipCode = mstiMwpDto.zipCode;

        }
    }
}
