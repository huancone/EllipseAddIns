using System;
using System.Collections.Generic;

namespace EllipseJobsClassLibrary
{


    public class Jobs
    {

        public string AccountCode { get; set; }
        public string ActDurHrs { get; set; }
        public string ActEquipCost { get; set; }
        public string ActLabCost { get; set; }
        public string ActMatCost { get; set; }
        public string ActOtherCost { get; set; }
        public string ActualFinishTime { get; set; }
        public string ActualStartDate { get; set; }
        public string ActualStartTime { get; set; }
        public string AptwExistsSw { get; set; }
        public string AssignPerson { get; set; }
        public string AssocEquipSw { get; set; }

        public string AssumeFirstMSTI { get; set; }
        //public JobLinkDTO autoGroupProjection { get; set; }

        public string CalcEquipCost { get; set; }
        public string CalcLabCost { get; set; }
        public string CalcLabHrs { get; set; }
        public string CalcMatCost { get; set; }

        public string CalcOthCost { get; set; }

        //public JobsJobLinkDTO[] childLinks { get; set; }
        public string ClosedDt { get; set; }
        public string ClosedTime { get; set; }
        public string CompCode { get; set; }
        public string CompModCode { get; set; }
        public string CompletedCode { get; set; }
        public string ConAstSegFr { get; set; }
        public string ConAstSegTo { get; set; }
        public string CountyShire { get; set; }
        public string Crew { get; set; }
        public string CrteInsitu { get; set; }

        public string Data1732 { get; set; }
        public string DateStatus { get; set; }
        public string DstrctAcctCode { get; set; }
        public string DstrctCode { get; set; }

        public string EmailAddress { get; set; }
        public string EquipClass { get; set; }
        public string EquipClassifx1 { get; set; }
        public string EquipClassifx10 { get; set; }
        public string EquipClassifx11 { get; set; }
        public string EquipClassifx12 { get; set; }
        public string EquipClassifx13 { get; set; }
        public string EquipClassifx14 { get; set; }
        public string EquipClassifx15 { get; set; }
        public string EquipClassifx16 { get; set; }
        public string EquipClassifx17 { get; set; }
        public string EquipClassifx18 { get; set; }
        public string EquipClassifx19 { get; set; }
        public string EquipClassifx2 { get; set; }
        public string EquipClassifx20 { get; set; }
        public string EquipClassifx3 { get; set; }
        public string EquipClassifx4 { get; set; }
        public string EquipClassifx5 { get; set; }
        public string EquipClassifx6 { get; set; }
        public string EquipClassifx7 { get; set; }
        public string EquipClassifx8 { get; set; }
        public string EquipClassifx9 { get; set; }
        public string EquipGrpId { get; set; }
        public string EquipLocation { get; set; }
        public string EquipNo { get; set; }
        public string EquipStatus { get; set; }
        public string EquipUpdateFlag { get; set; }
        public string EstDurHrs { get; set; }
        public string EstEquipCost { get; set; }
        public string EstLabCost { get; set; }
        public string EstLabHrs { get; set; }
        public string EstMatCost { get; set; }
        public string EstOtherCost { get; set; }

        public string FailurePart { get; set; }
        public string FaxNumber { get; set; }
        public string FromLink { get; set; }

        public string GanttFinishDateTime { get; set; }
        public string GanttLinkId { get; set; }
        public string GanttParentLinkId { get; set; }
        public string GanttStartDateTime { get; set; }

        public string HasAssignTodo { get; set; }
        public string Id { get; set; }
        public string IsControl { get; set; }
        public string IsMSTParent { get; set; }
        public string IsParent { get; set; }
        public string ItemName1 { get; set; }
        public string ItemName2 { get; set; }
        public string JobId { get; set; }
        public string JobParentId { get; set; }
        public string JobType { get; set; }
        public string LastPerformedDate { get; set; }
        public string LastPerformedDstrctCode { get; set; }
        public string LastPerformedWorkOrder { get; set; }
        public string LinkDistrictCode { get; set; }
        public string LinkEntityId { get; set; }
        public string LinkEntityType { get; set; }
        public string LinkErpRef { get; set; }
        public string LinkErpTaskRef { get; set; }
        public string LinkId { get; set; }
        public string LinkJobLinkListId { get; set; }
        public string LinkLagOrLead { get; set; }
        public string LinkOffset { get; set; }
        public string LinkScale { get; set; }
        public string LinkScaledOffset { get; set; }
        public string LinkSequenceNum { get; set; }
        public string LinkStatType { get; set; }
        public string LinkStatValue { get; set; }
        public string LinkState { get; set; }
        public string LinkType { get; set; }

        //public JobsMWPDTO[] linkedChildren { get; set; }
        public string LinkedInd { get; set; }
        public string Location { get; set; }
        public string LocationFr { get; set; }

        public string MaintSchTask { get; set; }
        public string MaintType { get; set; }
        public string MatUpdateFlag { get; set; }
        public string Mnemonic { get; set; }
        public string MstReference { get; set; }
        public string MultipleResourceRequirements { get; set; }

        public string NoGanttMove { get; set; }
        public string NoGanttResize { get; set; }
        public string NoOfTasks { get; set; }
        public string NumberOfMSTisRemaining { get; set; }

        public string OrigPriority { get; set; }
        public string OriginalPlannedFinishDate { get; set; }
        public string OriginalPlannedFinishTime { get; set; }
        public string OriginalPlannedStartDate { get; set; }
        public string OriginalPlannedStartTime { get; set; }
        public string OriginatorId { get; set; }
        public string OtherUpdateFlag { get; set; }
        public string ParentEntityId { get; set; }
        public string ParentEquip { get; set; }
        public string ParentJobType { get; set; }
        public string ParentLinkId { get; set; }
        public string ParentPlantNo { get; set; }
        public string ParentWo { get; set; }
        public string PartNo { get; set; }
        public string PartialCacheKey { get; set; }
        public string PcComplete { get; set; }
        public string PlanFinDate { get; set; }
        public string PlanFinTime { get; set; }
        public string PlanPriority { get; set; }
        public string PlanStatType { get; set; }
        public string PlanStatVal { get; set; }
        public string PlanStrDate { get; set; }
        public string PlanStrTime { get; set; }
        public string PlantNo { get; set; }
        public string PrefDate { get; set; }
        public string PrefTime { get; set; }
        public string PrinterName { get; set; }
        public string ProdUnitItem { get; set; }
        public string ProjDesc { get; set; }
        public string ProjectNo { get; set; }

        public string RaisedDate { get; set; }
        public string RaisedTime { get; set; }
        public string RecallTimeHrs { get; set; }
        public string Reference { get; set; }
        public string ReqByDate { get; set; }
        public string ReqByTime { get; set; }
        public string ReqStartDate { get; set; }
        public string ReqStartTime { get; set; }
        public string RequestId { get; set; }
        public string ResUpdateFlag { get; set; }

        //public JobsResourceDTO[] resourceRequirements { get; set; }

        public string RestartChildWODstrctCode { get; set; }
        public string RestartChildWOPlanStrDate { get; set; }
        public string RestartChildWOWorkOrder { get; set; }
        public string RestartMSTIFromLinkMstReference { get; set; }
        public string RestartMSTIFromLinkPlanStrDate { get; set; }
        public string RestartMSTIMstReference { get; set; }
        public string RestartMSTIPlanStrDate { get; set; }
        public string RestartParentWODstrctCode { get; set; }
        public string RestartParentWOPlanStrDate { get; set; }
        public string RestartParentWOWorkOrder { get; set; }

        //public TaskKeyDTO[] restartPrevLinks { get; set; }

        public string SchSegFr { get; set; }
        public string SchSegTo { get; set; }
        public string SchedDesc2 { get; set; }
        public string SequenceNo { get; set; }
        public string SerialNumber { get; set; }
        public string ShortDesc1 { get; set; }
        public string ShortDesc2 { get; set; }
        public string ShutdownNo { get; set; }
        public string ShutdownType { get; set; }
        public string Source { get; set; }
        public string State { get; set; }
        public string StatutoryFlg { get; set; }
        public string StdJobNo { get; set; }
        public string StreetName { get; set; }
        public string StreetNo { get; set; }
        public string Suburb { get; set; }
        public string SuppressingMST { get; set; }

        public string TargtFinDate { get; set; }
        public string TargtStrDate { get; set; }
        public string TaskAptwSw { get; set; }
        public string TownCity { get; set; }

        public string UnitOfWork { get; set; }
        public string UnitsComplete { get; set; }
        public string UnitsRequired { get; set; }

        public string WoDesc { get; set; }
        public string WoJobCodex1 { get; set; }
        public string WoJobCodex10 { get; set; }
        public string WoJobCodex2 { get; set; }
        public string WoJobCodex3 { get; set; }
        public string WoJobCodex4 { get; set; }
        public string WoJobCodex5 { get; set; }
        public string WoJobCodex6 { get; set; }
        public string WoJobCodex7 { get; set; }
        public string WoJobCodex8 { get; set; }
        public string WoJobCodex9 { get; set; }
        public string WoStatusM { get; set; }
        public string WoStatusU { get; set; }
        public string WoTaskNo { get; set; }
        public string WoType { get; set; }
        public string WorkGroup { get; set; }
        public string WorkOrder { get; set; }
        public string WorkRequestDescription { get; set; }

        public string ZipCode { get; set; }

        public Jobs()
        {

        }

        public Jobs(JobsMWPService.JobsMWPDTO jobDto)
        {
            AccountCode = jobDto.accountCode;
            ActDurHrs = jobDto.actDurHrsSpecified ? "" + jobDto.actDurHrs : null;
            ActEquipCost = jobDto.actEquipCostSpecified ? "" + jobDto.actEquipCost : null;
            ActLabCost = jobDto.actLabCostSpecified ? "" + jobDto.actLabCost : null;
            ActMatCost = jobDto.actMatCostSpecified ? "" + jobDto.actMatCost : null;
            ActOtherCost = jobDto.actOtherCostSpecified ? "" + jobDto.actOtherCost : null;
            ActualFinishTime = jobDto.actualFinishTime;
            ActualStartDate = jobDto.actualStartDateSpecified ? "" + jobDto.actualStartDate : null;
            ActualStartTime = jobDto.actualStartTime;
            AptwExistsSw = jobDto.aptwExistsSwSpecified ? "" + jobDto.aptwExistsSw : null;
            AssignPerson = jobDto.assignPerson;
            AssocEquipSw = jobDto.assocEquipSwSpecified ? "" + jobDto.assocEquipSw : null;
            AssumeFirstMSTI = jobDto.assumeFirstMSTISpecified ? "" + jobDto.assumeFirstMSTI : null;
            //public JobLinkDTO autoGroupProjection = jobDto.autoGroupProjection;
            
            CalcEquipCost = jobDto.calcEquipCostSpecified ? "" + jobDto.calcEquipCost : null;
            CalcLabCost = jobDto.calcLabCostSpecified ? "" + jobDto.calcLabCost : null;
            CalcLabHrs = jobDto.calcLabHrsSpecified ? "" + jobDto.calcLabHrs : null;
            CalcMatCost = jobDto.calcMatCostSpecified ? "" + jobDto.calcMatCost : null;
            CalcOthCost = jobDto.calcOthCostSpecified ? "" + jobDto.calcOthCost : null;
            //public JobsJobLinkDTO[] childLinks = jobDto.childLinks;
            ClosedDt = jobDto.closedDtSpecified ? "" + jobDto.closedDt : null;
            ClosedTime = jobDto.closedTimeSpecified ? "" + jobDto.closedTime : null;
            CompCode = jobDto.compCode;
            CompModCode = jobDto.compModCode;
            CompletedCode = jobDto.completedCode;
            ConAstSegFr = jobDto.conAstSegFrSpecified ? "" + jobDto.conAstSegFr : null;
            ConAstSegTo = jobDto.conAstSegToSpecified ? "" + jobDto.conAstSegTo : null;
            CountyShire = jobDto.countyShire;
            Crew = jobDto.crew;
            CrteInsitu = jobDto.crteInsituSpecified ? "" + jobDto.crteInsitu : null;

            Data1732 = jobDto.data1732;
            DateStatus = jobDto.dateStatus;
            DstrctAcctCode = jobDto.dstrctAcctCode;
            DstrctCode = jobDto.dstrctCode;

            EmailAddress = jobDto.emailAddress;
            EquipClass = jobDto.equipClass;
            EquipClassifx1 = jobDto.equipClassifx1;
            EquipClassifx10 = jobDto.equipClassifx10;
            EquipClassifx11 = jobDto.equipClassifx11;
            EquipClassifx12 = jobDto.equipClassifx12;
            EquipClassifx13 = jobDto.equipClassifx13;
            EquipClassifx14 = jobDto.equipClassifx14;
            EquipClassifx15 = jobDto.equipClassifx15;
            EquipClassifx16 = jobDto.equipClassifx16;
            EquipClassifx17 = jobDto.equipClassifx17;
            EquipClassifx18 = jobDto.equipClassifx18;
            EquipClassifx19 = jobDto.equipClassifx19;
            EquipClassifx2 = jobDto.equipClassifx2;
            EquipClassifx20 = jobDto.equipClassifx20;
            EquipClassifx3 = jobDto.equipClassifx3;
            EquipClassifx4 = jobDto.equipClassifx4;
            EquipClassifx5 = jobDto.equipClassifx5;
            EquipClassifx6 = jobDto.equipClassifx6;
            EquipClassifx7 = jobDto.equipClassifx7;
            EquipClassifx8 = jobDto.equipClassifx8;
            EquipClassifx9 = jobDto.equipClassifx9;
            EquipGrpId = jobDto.equipGrpId;
            EquipLocation = jobDto.equipLocation;
            EquipNo = jobDto.equipNo;
            EquipStatus = jobDto.equipStatus;
            EquipUpdateFlag = jobDto.equipUpdateFlagSpecified ? "" + jobDto.equipUpdateFlag : null;
            EstDurHrs = jobDto.estDurHrsSpecified ? "" + jobDto.estDurHrs : null;
            EstEquipCost = jobDto.estEquipCostSpecified ? "" + jobDto.estEquipCost : null;
            EstLabCost = jobDto.estLabCostSpecified ? "" + jobDto.estLabCost : null;
            EstLabHrs = jobDto.estLabHrsSpecified ? "" + jobDto.estLabHrs : null;
            EstMatCost = jobDto.estMatCostSpecified ? "" + jobDto.estMatCost : null;
            EstOtherCost = jobDto.estOtherCostSpecified ? "" + jobDto.estOtherCost : null;

            FailurePart = jobDto.failurePart;
            FaxNumber = jobDto.faxNumber;
            FromLink = jobDto.fromLinkSpecified ? "" + jobDto.fromLink : null;

            GanttFinishDateTime = jobDto.ganttFinishDateTime;
            GanttLinkId = jobDto.ganttLinkId;
            GanttParentLinkId = jobDto.ganttParentLinkId;
            GanttStartDateTime = jobDto.ganttStartDateTime;

            HasAssignTodo = jobDto.hasAssignTodo;
            Id = jobDto.idSpecified ? "" + jobDto.id : null;
            IsControl = jobDto.isControlSpecified ? "" + jobDto.isControl : null;
            IsMSTParent = jobDto.isMSTParent;
            IsParent = jobDto.isParent;
            ItemName1 = jobDto.itemName1;
            ItemName2 = jobDto.itemName2;
            JobId = jobDto.jobId;
            JobParentId = jobDto.jobParentId;
            JobType = jobDto.jobType;
            LastPerformedDate = jobDto.lastPerformedDateSpecified ? "" + jobDto.lastPerformedDate : null;
            LastPerformedDstrctCode = jobDto.lastPerformedDstrctCode;
            LastPerformedWorkOrder = jobDto.lastPerformedWorkOrder;
            LinkDistrictCode = jobDto.linkDistrictCode;
            LinkEntityId = jobDto.linkEntityIdSpecified ? "" + jobDto.linkEntityId : null;
            LinkEntityType = jobDto.linkEntityTypeSpecified ? "" + jobDto.linkEntityType : null;
            LinkErpRef = jobDto.linkErpRef;
            LinkErpTaskRef = jobDto.linkErpTaskRef;
            LinkId = jobDto.linkIdSpecified ? "" + jobDto.linkId : null;
            LinkJobLinkListId = jobDto.linkJobLinkListIdSpecified ? "" + jobDto.linkJobLinkListId : null;
            LinkLagOrLead = jobDto.linkLagOrLead;
            LinkOffset = jobDto.linkOffsetSpecified ? "" + jobDto.linkOffset : null;
            LinkScale = jobDto.linkScale;
            LinkScaledOffset = jobDto.linkScaledOffsetSpecified ? "" + jobDto.linkScaledOffset : null;
            LinkSequenceNum = jobDto.linkSequenceNumSpecified ? "" + jobDto.linkSequenceNum : null;
            LinkStatType = jobDto.linkStatType;
            LinkStatValue = jobDto.linkStatValueSpecified ? "" + jobDto.linkStatValue : null;
            LinkState = jobDto.linkState;
            LinkType = jobDto.linkType;

            //public JobsMWPDTO[] linkedChildren = jobDto.linkedChildren;
            LinkedInd = jobDto.linkedInd;
            Location = jobDto.location;
            LocationFr = jobDto.locationFr;

            MaintSchTask = jobDto.maintSchTask;
            MaintType = jobDto.maintType;
            MatUpdateFlag = jobDto.matUpdateFlagSpecified ? "" + jobDto.matUpdateFlag : null;
            Mnemonic = jobDto.mnemonic;
            MstReference = jobDto.mstReference;
            MultipleResourceRequirements = jobDto.multipleResourceRequirements;

            NoGanttMove = jobDto.noGanttMove;
            NoGanttResize = jobDto.noGanttResize;
            NoOfTasks = jobDto.noOfTasks;
            NumberOfMSTisRemaining = jobDto.numberOfMSTisRemainingSpecified ? "" + jobDto.numberOfMSTisRemaining : null;

            OrigPriority = jobDto.origPriority;
            OriginalPlannedFinishDate = jobDto.originalPlannedFinishDateSpecified ? "" + jobDto.originalPlannedFinishDate : null;
            OriginalPlannedFinishTime = jobDto.originalPlannedFinishTime;
            OriginalPlannedStartDate = jobDto.originalPlannedStartDateSpecified ? "" + jobDto.originalPlannedStartDate : null;
            OriginalPlannedStartTime = jobDto.originalPlannedStartTime;
            OriginatorId = jobDto.originatorId;
            OtherUpdateFlag = jobDto.otherUpdateFlagSpecified ? "" + jobDto.otherUpdateFlag : null;
            ParentEntityId = jobDto.parentEntityIdSpecified ? "" + jobDto.parentEntityId : null;
            ParentEquip = jobDto.parentEquip;
            ParentJobType = jobDto.parentJobType;
            ParentLinkId = jobDto.parentLinkIdSpecified ? "" + jobDto.parentLinkId : null;
            ParentPlantNo = jobDto.parentPlantNo;
            ParentWo = jobDto.parentWo;
            PartNo = jobDto.partNo;
            PartialCacheKey = jobDto.partialCacheKey;
            PcComplete = jobDto.pcCompleteSpecified ? "" + jobDto.pcComplete : null;
            PlanFinDate = jobDto.planFinDateSpecified ? "" + jobDto.planFinDate : null;
            PlanFinTime = jobDto.planFinTime;
            PlanPriority = jobDto.planPriority;
            PlanStatType = jobDto.planStatType;
            PlanStatVal = jobDto.planStatValSpecified ? "" + jobDto.planStatVal : null;
            PlanStrDate = jobDto.planStrDateSpecified ? "" + jobDto.planStrDate : null;
            PlanStrTime = jobDto.planStrTime;
            PlantNo = jobDto.plantNo;
            PrefDate = jobDto.prefDateSpecified ? "" + jobDto.prefDate : null;
            PrefTime = jobDto.prefTime;
            PrinterName = jobDto.printerName;
            ProdUnitItem = jobDto.prodUnitItemSpecified ? "" + jobDto.prodUnitItem : null;
            ProjDesc = jobDto.projDesc;
            ProjectNo = jobDto.projectNo;

            RaisedDate = jobDto.raisedDateSpecified ? "" + jobDto.raisedDate : null;
            RaisedTime = jobDto.raisedTimeSpecified ? "" + jobDto.raisedTime : null;
            RecallTimeHrs = jobDto.recallTimeHrsSpecified ? "" + jobDto.recallTimeHrs : null;
            Reference = jobDto.reference;
            ReqByDate = jobDto.reqByDateSpecified ? "" + jobDto.reqByDate : null;
            ReqByTime = jobDto.reqByTime;
            ReqStartDate = jobDto.reqStartDateSpecified ? "" + jobDto.reqStartDate : null;
            ReqStartTime = jobDto.reqStartTime;
            RequestId = jobDto.requestId;
            ResUpdateFlag = jobDto.resUpdateFlagSpecified ? "" + jobDto.resUpdateFlag : null;

            //public JobsResourceDTO[] resourceRequirements = jobDto.resourceRequirements;

            RestartChildWODstrctCode = jobDto.restartChildWODstrctCode;
            RestartChildWOPlanStrDate = jobDto.restartChildWOPlanStrDateSpecified ? "" + jobDto.restartChildWOPlanStrDate : null;
            RestartChildWOWorkOrder = jobDto.restartChildWOWorkOrder;
            RestartMSTIFromLinkMstReference = jobDto.restartMSTIFromLinkMstReference;
            RestartMSTIFromLinkPlanStrDate = jobDto.restartMSTIFromLinkPlanStrDateSpecified ? "" + jobDto.restartMSTIFromLinkPlanStrDate : null;
            RestartMSTIMstReference = jobDto.restartMSTIMstReference;
            RestartMSTIPlanStrDate = jobDto.restartMSTIPlanStrDateSpecified ? "" + jobDto.restartMSTIPlanStrDate : null;
            RestartParentWODstrctCode = jobDto.restartParentWODstrctCode;
            RestartParentWOPlanStrDate = jobDto.restartParentWOPlanStrDateSpecified ? "" + jobDto.restartParentWOPlanStrDate : null;
            RestartParentWOWorkOrder = jobDto.restartParentWOWorkOrder;

            //public TaskKeyDTO[] restartPrevLinks = jobDto.xx;

            SchSegFr = jobDto.schSegFrSpecified ? "" + jobDto.schSegFr : null;
            SchSegTo = jobDto.schSegToSpecified ? "" + jobDto.schSegTo : null;
            SchedDesc2 = jobDto.schedDesc2;
            SequenceNo = jobDto.sequenceNoSpecified ? "" + jobDto.sequenceNo : null;
            SerialNumber = jobDto.serialNumber;
            ShortDesc1 = jobDto.shortDesc1;
            ShortDesc2 = jobDto.shortDesc2;
            ShutdownNo = jobDto.shutdownNo;
            ShutdownType = jobDto.shutdownType;
            Source = jobDto.source;
            State = jobDto.state;
            StatutoryFlg = jobDto.statutoryFlg;
            StdJobNo = jobDto.stdJobNo;
            StreetName = jobDto.streetName;
            StreetNo = jobDto.streetNo;
            Suburb = jobDto.suburb;
            SuppressingMST = jobDto.suppressingMST;

            TargtFinDate = jobDto.targtFinDateSpecified ? "" + jobDto.targtFinDate : null;
            TargtStrDate = jobDto.targtStrDateSpecified ? "" + jobDto.targtStrDate : null;
            TaskAptwSw = jobDto.taskAptwSw;
            TownCity = jobDto.townCity;

            UnitOfWork = jobDto.unitOfWork;
            UnitsComplete = jobDto.unitsCompleteSpecified ? "" + jobDto.unitsComplete : null;
            UnitsRequired = jobDto.unitsRequiredSpecified ? "" + jobDto.unitsRequired : null;

            WoDesc = jobDto.woDesc;
            WoJobCodex1 = jobDto.woJobCodex1;
            WoJobCodex10 = jobDto.woJobCodex10;
            WoJobCodex2 = jobDto.woJobCodex2;
            WoJobCodex3 = jobDto.woJobCodex3;
            WoJobCodex4 = jobDto.woJobCodex4;
            WoJobCodex5 = jobDto.woJobCodex5;
            WoJobCodex6 = jobDto.woJobCodex6;
            WoJobCodex7 = jobDto.woJobCodex7;
            WoJobCodex8 = jobDto.woJobCodex8;
            WoJobCodex9 = jobDto.woJobCodex9;
            WoStatusM = jobDto.woStatusM;
            WoStatusU = jobDto.woStatusU;
            WoTaskNo = jobDto.WoTaskNo;
            WoType = jobDto.woType;
            WorkGroup = jobDto.workGroup;
            WorkOrder = jobDto.workOrder;
            WorkRequestDescription = jobDto.workRequestDescription;
        }
    }



}
