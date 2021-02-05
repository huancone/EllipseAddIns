using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMsssEquipmentExcelAddIn
{
    public class MssItemDto
    {
        public string EquipmentGrpId;
        public string CompCode;
        public string CompModCode;
        public string CompCodeDescription;//
        public string FailureMode;
        public string FailureModeDescription;//
        public string FailureCode;
        public string FailureCodeDescription;//
        public string FunctionCode;
        public string FunctionCodeDescription;//
        public string Consequence;
        public string ConsequenceDescription;//
        public string Effect;
        public string Strategy;
        public string StrategyDescription;
        public string AgreedAction;
        public string FailureClass;
        public string FailureClassDescription;//
        public string FunctionClass;
        public string FunctionClassDescription;//

        public MSSSService.MSSSServiceCreateRequestDTO ToCreateRequestDto()
        {
            var item = new MSSSService.MSSSServiceCreateRequestDTO();
            
            item.equipmentGrpId = EquipmentGrpId;
            item.compCode = CompCode;
            item.compModCode = CompModCode;
            item.failureMode = FailureMode;
            item.failureCode = FailureCode;
            item.functionCode = FunctionCode;
            item.consequence = Consequence;
            item.effect = Effect;
            item.strategy = Strategy;
            item.agreedAction = AgreedAction;
            item.failureClass = FailureClass;
            item.functionClass = FunctionClass;

            return item;
        }

        public MSSSService.MSSSServiceDeleteRequestDTO ToDeleteRequestDto()
        {
            var item = new MSSSService.MSSSServiceDeleteRequestDTO();

            item.equipmentGrpId = EquipmentGrpId;
            item.compCode = CompCode;
            item.compModCode = CompModCode;
            item.failureMode = FailureMode;
            item.failureCode = FailureCode;
            item.functionCode = FunctionCode;
            item.consequence = Consequence;

            return item;
        }
    }
}
