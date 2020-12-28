using SharedClassLibrary.Utilities;

namespace EllipseJobsClassLibrary
{


    public class Jobs : SharedClassLibrary.Ellipse.PrimitiveClasses.JobsMwpDto
    {

        public Jobs(JobsMWPService.JobsMWPDTO jobDto)
        {
            AccountCode = jobDto.accountCode;
            ActDurHrs = jobDto.actDurHrsSpecified ? "" + jobDto.actDurHrs : null;
            ActEquipCost = jobDto.actEquipCostSpecified ? "" + jobDto.actEquipCost : null;
            ActLabCost = jobDto.actLabCostSpecified ? "" + jobDto.actLabCost : null;
            ActMatCost = jobDto.actMatCostSpecified ? "" + jobDto.actMatCost : null;
            ActOtherCost = jobDto.actOtherCostSpecified ? "" + jobDto.actOtherCost : null;
            ActualFinishDate = jobDto.actualFinishDateSpecified ? "" + MyUtilities.ToString(jobDto.actualFinishDate) : null;
            ActualFinishTime = jobDto.actualFinishTime;
            ActualStartDate = jobDto.actualStartDateSpecified ? "" + MyUtilities.ToString(jobDto.actualStartDate) : null;
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
            ClosedDt = jobDto.closedDtSpecified ? "" + MyUtilities.ToString(jobDto.closedDt) : null;
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
            LastPerformedDate = jobDto.lastPerformedDateSpecified ? "" + MyUtilities.ToString(jobDto.lastPerformedDate) : null;
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
            OriginalPlannedFinishDate = jobDto.originalPlannedFinishDateSpecified ? "" + MyUtilities.ToString(jobDto.originalPlannedFinishDate) : null;
            OriginalPlannedFinishTime = jobDto.originalPlannedFinishTime;
            OriginalPlannedStartDate = jobDto.originalPlannedStartDateSpecified ? "" + MyUtilities.ToString(jobDto.originalPlannedStartDate) : null;
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
            PlanFinDate = jobDto.planFinDateSpecified ? "" + MyUtilities.ToString(jobDto.planFinDate) : null;
            PlanFinTime = jobDto.planFinTime;
            PlanPriority = jobDto.planPriority;
            PlanStatType = jobDto.planStatType;
            PlanStatVal = jobDto.planStatValSpecified ? "" + jobDto.planStatVal : null;
            PlanStrDate = jobDto.planStrDateSpecified ? "" + MyUtilities.ToString(jobDto.planStrDate) : null;
            PlanStrTime = jobDto.planStrTime;
            PlantNo = jobDto.plantNo;
            PrefDate = jobDto.prefDateSpecified ? "" + MyUtilities.ToString(jobDto.prefDate) : null;
            PrefTime = jobDto.prefTime;
            PrinterName = jobDto.printerName;
            ProdUnitItem = jobDto.prodUnitItemSpecified ? "" + jobDto.prodUnitItem : null;
            ProjDesc = jobDto.projDesc;
            ProjectNo = jobDto.projectNo;

            RaisedDate = jobDto.raisedDateSpecified ? "" + MyUtilities.ToString(jobDto.raisedDate) : null;
            RaisedTime = jobDto.raisedTimeSpecified ? "" + jobDto.raisedTime : null;
            RecallTimeHrs = jobDto.recallTimeHrsSpecified ? "" + jobDto.recallTimeHrs : null;
            Reference = jobDto.reference;
            ReqByDate = jobDto.reqByDateSpecified ? "" + MyUtilities.ToString(jobDto.reqByDate) : null;
            ReqByTime = jobDto.reqByTime;
            ReqStartDate = jobDto.reqStartDateSpecified ? "" + MyUtilities.ToString(jobDto.reqStartDate) : null;
            ReqStartTime = jobDto.reqStartTime;
            RequestId = jobDto.requestId;
            ResUpdateFlag = jobDto.resUpdateFlagSpecified ? "" + jobDto.resUpdateFlag : null;

            //public JobsResourceDTO[] resourceRequirements = jobDto.resourceRequirements;

            RestartChildWODstrctCode = jobDto.restartChildWODstrctCode;
            RestartChildWOPlanStrDate = jobDto.restartChildWOPlanStrDateSpecified ? "" + MyUtilities.ToString(jobDto.restartChildWOPlanStrDate) : null;
            RestartChildWOWorkOrder = jobDto.restartChildWOWorkOrder;
            RestartMSTIFromLinkMstReference = jobDto.restartMSTIFromLinkMstReference;
            RestartMSTIFromLinkPlanStrDate = jobDto.restartMSTIFromLinkPlanStrDateSpecified ? "" + MyUtilities.ToString(jobDto.restartMSTIFromLinkPlanStrDate) : null;
            RestartMSTIMstReference = jobDto.restartMSTIMstReference;
            RestartMSTIPlanStrDate = jobDto.restartMSTIPlanStrDateSpecified ? "" + MyUtilities.ToString(jobDto.restartMSTIPlanStrDate) : null;
            RestartParentWODstrctCode = jobDto.restartParentWODstrctCode;
            RestartParentWOPlanStrDate = jobDto.restartParentWOPlanStrDateSpecified ? "" + MyUtilities.ToString(jobDto.restartParentWOPlanStrDate) : null;
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

            TargtFinDate = jobDto.targtFinDateSpecified ? "" + MyUtilities.ToString(jobDto.targtFinDate) : null;
            TargtStrDate = jobDto.targtStrDateSpecified ? "" + MyUtilities.ToString(jobDto.targtStrDate) : null;
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
