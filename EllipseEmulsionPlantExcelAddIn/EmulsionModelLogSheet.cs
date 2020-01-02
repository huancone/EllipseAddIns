using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseEmulsionPlantExcelAddIn.LogSheet;

namespace EllipseEmulsionPlantExcelAddIn
{
    public class EmulsionModelLogSheet
    {
        public string ModelName = "EMULPLANT";
        public string BlenderEquipmentNumber = "8210018";
        public string Silo1EquipmentNumber = "8210051";
        public string Silo2EquipmentNumber = "8210052";
        public string Silo3EquipmentNumber = "8210060";
        public string Date;
        public string ShiftCode;
        public string Operator;
        public string TotalEmStartValue;
        public string TotalEmEndValue;
        public string TotalEmProduced;
        public string TotalEmDispatched;
        public string TotalDestination;
        public string Silo1StartValue;
        public string Silo1EndValue;
        public string Silo1Produced;
        public string Silo2StartValue;
        public string Silo2EndValue;
        public string Silo2Produced;
        public string Silo3StartValue;
        public string Silo3EndValue;
        public string Silo3Produced;

        public LogSheetItem ToLogSheet()
        {
            var logSheet = new LogSheetItem(ModelName, Date, ShiftCode);

            logSheet.ModelName = ModelName;

            //Obtengo los encabezados adicionales MSO430
            var header = new List<LogSheetHeaderItem>();
            header.Add(new LogSheetHeaderItem("EMULPLANT", 1, "TNPROD", "TN"));
            header.Add(new LogSheetHeaderItem("EMULPLANT", 2, "DESTIN", "SD"));
            header.Add(new LogSheetHeaderItem("EMULPLANT", 3, "TNDESP", "TT"));
            header.Add(new LogSheetHeaderItem("EMULPLANT", 4, "MINIC", "MI"));
            header.Add(new LogSheetHeaderItem("EMULPLANT", 5, "MFINAL", "MF"));

            //Establezco el listado de equipos del modelo MSO460 basado en el MSO615. (Índices no directamente relacionado con los inputs del modelo)
            #region ItemModel
            var equipments = new List<LogSheetEquipmentModelItem>();

            var eq01 = new LogSheetEquipmentModelItem
            {
                EntryGrp = BlenderEquipmentNumber,

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
                DestinationLocation = Silo1EquipmentNumber,
                MaterialCode = "EM",

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
                EntryGrp = Silo1EquipmentNumber,

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
                DestinationLocation = Silo1EquipmentNumber,
                MaterialCode = "EM",

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
                EntryGrp = Silo2EquipmentNumber,

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
                DestinationLocation = Silo2EquipmentNumber,
                MaterialCode = "EM",

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
                EntryGrp = Silo3EquipmentNumber,

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
                DestinationLocation = Silo3EquipmentNumber,
                MaterialCode = "EM",

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
            #endregion
            logSheet.ModelItems = equipments;

            #region Inputs
            var inputItems = new List<LogSheetEquipmentInputItem>();
            var blender = new LogSheetEquipmentInputItem();
            blender.PlantNo.Value = BlenderEquipmentNumber;
            blender.Operator.Value = Operator;
            blender.Input1.Value = TotalEmProduced;
            blender.Input2.Value = TotalDestination;
            blender.Input3.Value = TotalEmDispatched;
            blender.Input4.Value = TotalEmStartValue;
            blender.Input5.Value = TotalEmEndValue;
            inputItems.Add(blender);

            var silo1 = new LogSheetEquipmentInputItem();
            silo1.PlantNo.Value = Silo1EquipmentNumber;
            silo1.Operator.Value = Operator;
            silo1.Input1.Value = Silo1Produced;
            silo1.Input2.Value = Silo1EquipmentNumber;
            silo1.Input3.Value = null;
            silo1.Input4.Value = Silo1StartValue;
            silo1.Input5.Value = Silo1EndValue;
            inputItems.Add(silo1);

            var silo2 = new LogSheetEquipmentInputItem();
            silo2.PlantNo.Value = Silo2EquipmentNumber;
            silo2.Operator.Value = Operator;
            silo2.Input1.Value = Silo2Produced;
            silo2.Input2.Value = Silo2EquipmentNumber;
            silo2.Input3.Value = null;
            silo2.Input4.Value = Silo2StartValue;
            silo2.Input5.Value = Silo2EndValue;
            inputItems.Add(silo2);

            var silo3 = new LogSheetEquipmentInputItem();
            silo3.PlantNo.Value = Silo3EquipmentNumber;
            silo3.Operator.Value = Operator;
            silo3.Input1.Value = Silo3Produced;
            silo3.Input2.Value = Silo3EquipmentNumber;
            silo3.Input3.Value = null;
            silo3.Input4.Value = Silo3StartValue;
            silo3.Input5.Value = Silo3EndValue;
            inputItems.Add(silo3);
            #endregion
            logSheet.InputItems = inputItems;
            return logSheet;
        }
    }
}