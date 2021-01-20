using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseEmulsionPlantExcelAddIn.LogSheet
{
    public class LogSheetEquipmentModelItem
    {
        //Flags: I - Input / O - Output / B - Both Input and Output 
        public string ModelCode;
        public string Egi;
        public string ModelSeqNo;
        public string EntryGrp;
        //Flags: O Optional, M Mandatory, N Not Required
        public string OperatorFlag;
        public string AccountFlag;
        public string WorkOrderFlag;
        public string SourceLocationFlag;
        public string DestinationLocationFlag;
        public string MaterialFlag;
        //
        public string OperatorId;
        public string Account;
        public string WorkOrder;
        public string SourceLocation;
        public string DestinationLocation;
        public string MaterialCode;
        //Flags: I Input, O Output, B Both Input and Output From Model
        public string StatValue1Flag;
        public string StatValue2Flag;
        public string StatValue3Flag;
        public string StatValue4Flag;
        public string StatValue5Flag;
        public string StatValue6Flag;
        public string StatValue7Flag;
        public string StatValue8Flag;
        public string StatValue9Flag;
        public string StatValue10Flag;
        //
        public string StatValue1;
        public string StatValue2;
        public string StatValue3;
        public string StatValue4;
        public string StatValue5;
        public string StatValue6;
        public string StatValue7;
        public string StatValue8;
        public string StatValue9;
        public string StatValue10;
    }
}
