using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseEmulsionPlantExcelAddIn.LogSheet;

namespace EllipseEmulsionPlantExcelAddIn
{
    public class SolutionModelLogSheet
    {
        public string ModelName = "EMULPLANT2";
        public string RotatorEquipmentNumber = "8210102";
        public string Tank1EquipmentNumber = "8210016";
        public string Tank2EquipmentNumber = "8210030";
        public string Tank3EquipmentNumber = "8210031";
        public string Tank4EquipmentNumber = "8210032";
        public string Date;
        public string ShiftCode;
        public string Operator;
        public string TotalSolStartValue;
        public string TotalSolEndValue;
        public string TotalSolProduced;
        public string TotalSolUsed;
        public string TotalDestination;
        public string Tank1StartValue;
        public string Tank1EndValue;
        public string Tank1Produced;
        public string Tank2StartValue;
        public string Tank2EndValue;
        public string Tank2Produced;
        public string Tank3StartValue;
        public string Tank3EndValue;
        public string Tank3Produced;
        public string Tank4StartValue;
        public string Tank4EndValue;
        public string Tank4Produced;

        public LogSheetItem ToLogSheet()
        {
            var logSheet = new LogSheetItem(ModelName, Date, ShiftCode);

            logSheet.ModelName = ModelName;

            //Obtengo los encabezados adicionales MSO430
            var header = new List<LogSheetHeaderItem>();
            header.Add(new LogSheetHeaderItem("EMULPLANT2", 1, "TNPROD", "TN"));
            header.Add(new LogSheetHeaderItem("EMULPLANT2", 2, "DESTIN", "SD"));
            header.Add(new LogSheetHeaderItem("EMULPLANT2", 3, "TUSADA", "TT"));
            header.Add(new LogSheetHeaderItem("EMULPLANT2", 4, "MINIC", "MI"));
            header.Add(new LogSheetHeaderItem("EMULPLANT2", 5, "MFINAL", "MF"));

            //Establezco el listado de equipos del modelo MSO460 basado en el MSO615. (Índices no directamente relacionado con los inputs del modelo)
            #region ItemModel
            var equipments = new List<LogSheetEquipmentModelItem>();

            var eq01 = new LogSheetEquipmentModelItem
            {
                EntryGrp = RotatorEquipmentNumber,

                OperatorFlag = "O",
                AccountFlag = "N",
                WorkOrderFlag = "N",
                SourceLocationFlag = "N",
                DestinationLocationFlag = "O",
                MaterialFlag = "M",

                OperatorId = "",
                Account = "",
                WorkOrder = "",
                SourceLocation = "",
                DestinationLocation = Tank1EquipmentNumber,
                MaterialCode = "SOLOX",

                StatValue1Flag = "I",
                StatValue2Flag = "I",
                StatValue3Flag = "B",
                StatValue4Flag = "I",
                StatValue5Flag = "I",
                StatValue6Flag = null,
                StatValue7Flag = null,
                StatValue8Flag = null,
                StatValue9Flag = null,
                StatValue10Flag = null,

                StatValue1 = "0",
                StatValue2 = "0",
                StatValue3 = "0",
                StatValue4 = "0",
                StatValue5 = "0",
                StatValue6 = "0",
                StatValue7 = "0",
                StatValue8 = "0",
                StatValue9 = "0",
                StatValue10 = "0"
            };
            equipments.Add(eq01);

            var eq02 = new LogSheetEquipmentModelItem
            {
                EntryGrp = Tank1EquipmentNumber,

                OperatorFlag = "O",
                AccountFlag = "N",
                WorkOrderFlag = "N",
                SourceLocationFlag = "N",
                DestinationLocationFlag = "O",
                MaterialFlag = "M",

                OperatorId = "",
                Account = "",
                WorkOrder = "",
                SourceLocation = "",
                DestinationLocation = Tank1EquipmentNumber,
                MaterialCode = "SOLOX",

                StatValue1Flag = "I",
                StatValue2Flag = "I",
                StatValue3Flag = "B",
                StatValue4Flag = "I",
                StatValue5Flag = null,
                StatValue6Flag = null,
                StatValue7Flag = null,
                StatValue8Flag = null,
                StatValue9Flag = null,
                StatValue10Flag = null,

                StatValue1 = "0",
                StatValue2 = "0",
                StatValue3 = "0",
                StatValue4 = "0",
                StatValue5 = "0",
                StatValue6 = "0",
                StatValue7 = "0",
                StatValue8 = "0",
                StatValue9 = "0",
                StatValue10 = "0"
            };
            equipments.Add(eq02);

            var eq03 = new LogSheetEquipmentModelItem
            {
                EntryGrp = Tank2EquipmentNumber,

                OperatorFlag = "O",
                AccountFlag = "N",
                WorkOrderFlag = "N",
                SourceLocationFlag = "N",
                DestinationLocationFlag = "O",
                MaterialFlag = "M",

                OperatorId = "",
                Account = "",
                WorkOrder = "",
                SourceLocation = "",
                DestinationLocation = Tank2EquipmentNumber,
                MaterialCode = "SOLOX",

                StatValue1Flag = "I",
                StatValue2Flag = "I",
                StatValue3Flag = "B",
                StatValue4Flag = "I",
                StatValue5Flag = null,
                StatValue6Flag = null,
                StatValue7Flag = null,
                StatValue8Flag = null,
                StatValue9Flag = null,
                StatValue10Flag = null,

                StatValue1 = "0",
                StatValue2 = "0",
                StatValue3 = "0",
                StatValue4 = "0",
                StatValue5 = "0",
                StatValue6 = "0",
                StatValue7 = "0",
                StatValue8 = "0",
                StatValue9 = "0",
                StatValue10 = "0"
            };
            equipments.Add(eq03);

            var eq04 = new LogSheetEquipmentModelItem
            {
                EntryGrp = Tank3EquipmentNumber,

                OperatorFlag = "O",
                AccountFlag = "N",
                WorkOrderFlag = "N",
                SourceLocationFlag = "N",
                DestinationLocationFlag = "O",
                MaterialFlag = "M",

                OperatorId = "",
                Account = "",
                WorkOrder = "",
                SourceLocation = "",
                DestinationLocation = Tank3EquipmentNumber,
                MaterialCode = "SOLOX",

                StatValue1Flag = "I",
                StatValue2Flag = "I",
                StatValue3Flag = "B",
                StatValue4Flag = "I",
                StatValue5Flag = null,
                StatValue6Flag = null,
                StatValue7Flag = null,
                StatValue8Flag = null,
                StatValue9Flag = null,
                StatValue10Flag = null,

                StatValue1 = "0",
                StatValue2 = "0",
                StatValue3 = "0",
                StatValue4 = "0",
                StatValue5 = "0",
                StatValue6 = "0",
                StatValue7 = "0",
                StatValue8 = "0",
                StatValue9 = "0",
                StatValue10 = "0"
            };
            equipments.Add(eq04);

            var eq05 = new LogSheetEquipmentModelItem
            {
                EntryGrp = Tank3EquipmentNumber,

                OperatorFlag = "O",
                AccountFlag = "N",
                WorkOrderFlag = "N",
                SourceLocationFlag = "N",
                DestinationLocationFlag = "O",
                MaterialFlag = "M",

                OperatorId = "",
                Account = "",
                WorkOrder = "",
                SourceLocation = "",
                DestinationLocation = Tank4EquipmentNumber,
                MaterialCode = "SOLOX",

                StatValue1Flag = "I",
                StatValue2Flag = "I",
                StatValue3Flag = "B",
                StatValue4Flag = "I",
                StatValue5Flag = null,
                StatValue6Flag = null,
                StatValue7Flag = null,
                StatValue8Flag = null,
                StatValue9Flag = null,
                StatValue10Flag = null,

                StatValue1 = "0",
                StatValue2 = "0",
                StatValue3 = "0",
                StatValue4 = "0",
                StatValue5 = "0",
                StatValue6 = "0",
                StatValue7 = "0",
                StatValue8 = "0",
                StatValue9 = "0",
                StatValue10 = "0"
            };
            equipments.Add(eq05);
            #endregion
            logSheet.ModelItems = equipments;

            #region Inputs
            var inputItems = new List<LogSheetEquipmentInputItem>();
            var rotator = new LogSheetEquipmentInputItem();
            rotator.PlantNo.Value = RotatorEquipmentNumber;
            rotator.Operator.Value = Operator;
            rotator.Input1.Value = TotalSolProduced;
            rotator.Input2.Value = TotalDestination;
            rotator.Input3.Value = TotalSolUsed;
            rotator.Input4.Value = TotalSolStartValue;
            rotator.Input5.Value = TotalSolEndValue;
            inputItems.Add(rotator);

            var tank1 = new LogSheetEquipmentInputItem();
            tank1.PlantNo.Value = Tank1EquipmentNumber;
            tank1.Operator.Value = Operator;
            tank1.Input1.Value = Tank1Produced;
            tank1.Input2.Value = Tank1EquipmentNumber;
            tank1.Input3.Value = null;
            tank1.Input4.Value = Tank1StartValue;
            tank1.Input5.Value = Tank1EndValue;
            inputItems.Add(tank1);

            var tank2 = new LogSheetEquipmentInputItem();
            tank2.PlantNo.Value = Tank2EquipmentNumber;
            tank2.Operator.Value = Operator;
            tank2.Input1.Value = Tank2Produced;
            tank2.Input2.Value = Tank2EquipmentNumber;
            tank2.Input3.Value = null;
            tank2.Input4.Value = Tank2StartValue;
            tank2.Input5.Value = Tank2EndValue;
            inputItems.Add(tank2);

            var tank3 = new LogSheetEquipmentInputItem();
            tank3.PlantNo.Value = Tank3EquipmentNumber;
            tank3.Operator.Value = Operator;
            tank3.Input1.Value = Tank3Produced;
            tank3.Input2.Value = Tank3EquipmentNumber;
            tank3.Input3.Value = null;
            tank3.Input4.Value = Tank3StartValue;
            tank3.Input5.Value = Tank3EndValue;
            inputItems.Add(tank3);

            var tank4 = new LogSheetEquipmentInputItem();
            tank4.PlantNo.Value = Tank4EquipmentNumber;
            tank4.Operator.Value = Operator;
            tank4.Input1.Value = Tank4Produced;
            tank4.Input2.Value = Tank4EquipmentNumber;
            tank4.Input3.Value = null;
            tank4.Input4.Value = Tank4StartValue;
            tank4.Input5.Value = Tank4EndValue;
            inputItems.Add(tank4);


            #endregion
            logSheet.InputItems = inputItems;
            return logSheet;
        }
    }
}