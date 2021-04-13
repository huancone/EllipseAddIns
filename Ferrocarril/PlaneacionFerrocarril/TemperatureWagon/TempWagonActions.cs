using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;

namespace PlaneacionFerrocarril.TemperatureWagon
{
    public static class TempWagonActions
    {
        public static void TransformLogToPlain(string file, ExcelStyleCells cells, int startingRow)
        {
            var currentRow = startingRow;

            var enumLines = File.ReadAllLines(file, Encoding.UTF8);

            //Buscamos fecha y el valor de inicio de la información
            var date = "";
            var startIndex = 0;
            for (var i = 0; i < enumLines.Length; i++)
            {
                if (enumLines[i].StartsWith("Direction") && enumLines[i + 1].StartsWith("Speed In/Out"))
                    date = enumLines[i].Substring(66);

                if (enumLines[i].StartsWith("Order") && enumLines[i - 1].StartsWith("Car"))
                {
                    startIndex = i + 2;
                    break;
                }
            }

            LogCar previousLogCar = null;
            var emptyOwnerList = new List<int>();
            LogCar namedLogCar = null;
            //Llenamos los valores de la hoja plana
            for (var i = startIndex; i < enumLines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(enumLines[i]))
                    continue;
                if (enumLines[i].StartsWith("ITAKA", StringComparison.InvariantCultureIgnoreCase) ||
                    enumLines[i].StartsWith("URIBIA", StringComparison.InvariantCultureIgnoreCase) ||
                    enumLines[i].StartsWith("COSINA", StringComparison.InvariantCultureIgnoreCase) ||
                    enumLines[i].StartsWith("ISHAMANA", StringComparison.InvariantCultureIgnoreCase))
                    break;

                var logCar = new LogCar(enumLines[i]);

                //Corregimos el número del vagón a la notación de ellipse
                if (logCar.Number.Length == 4)
                    logCar.Number = "10000" + logCar.Number.Substring(logCar.Number.Length - 2);
                else if (logCar.Number.Length != 0)
                    logCar.Number = "110" + logCar.Number.Substring(logCar.Number.Length - 4);

                if (previousLogCar != null)
                {
                    //llenamos los espacios vacíos del número de orden del vagón
                    if (string.IsNullOrWhiteSpace(logCar.Order))
                    {
                        logCar.Order = previousLogCar.Order;
                    }
                    //si cambió se llenan los vacíos correspondientes al mismo
                    else if (!logCar.Order.Equals(previousLogCar.Order, StringComparison.InvariantCultureIgnoreCase))
                    {
                        //Rellene los nombrados
                        if (namedLogCar != null)
                        {
                            foreach (var index in emptyOwnerList)
                            {
                                cells.GetCell(3, index).Value = namedLogCar.Owner;
                                cells.GetCell(4, index).Value = namedLogCar.Number;
                            }
                        }

                        namedLogCar = null;
                        emptyOwnerList.Clear();
                    }
                }


                if (!string.IsNullOrWhiteSpace(logCar.Owner))
                    namedLogCar = logCar;
                else
                    emptyOwnerList.Add(currentRow);


                cells.GetCell(1, currentRow).Value = MyUtilities.ToString(MyUtilities.ToDate(date, "MM-dd-yyyy"), MyUtilities.DateTime.DateYYYYMMDD);
                cells.GetCell(2, currentRow).Value = logCar.Order;
                cells.GetCell(3, currentRow).Value = logCar.Owner;
                cells.GetCell(4, currentRow).Value = logCar.Number;
                cells.GetCell(5, currentRow).Value = logCar.AxleNumber;
                //Empty Space   ,
                cells.GetCell(7, currentRow).Value = logCar.Spacing;
                cells.GetCell(8, currentRow).Value = logCar.Ch1;
                cells.GetCell(9, currentRow).Value = logCar.Ch2;
                cells.GetCell(10, currentRow).Value = logCar.Ch3;
                cells.GetCell(11, currentRow).Value = logCar.Ch4;
                cells.GetCell(12, currentRow).Value = logCar.Alarms;

                previousLogCar = logCar;
                currentRow++;
            }

            //Rellene los nombrados
            if (namedLogCar != null)
            {
                foreach (var index in emptyOwnerList)
                {
                    cells.GetCell(3, index).Value = namedLogCar.Owner;
                    cells.GetCell(4, index).Value = namedLogCar.Number;
                }
            }
        }
    }
}
