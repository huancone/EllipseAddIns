using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseMaintSchedTaskClassLibrary
{
    public class MstForecast
    {
        public string Ninstances;

        public string CompCode;

        public string CompModCode;

        public string EquipNo;

        public string HideSuppressed;

        public string MaintSchTask;

        public string Rec700Type;

        public string ShowRelated;

        public MstForecast()
        {

        }
        public MstForecast(MstService.MSTForecastDTO mstForecastDto)
        {

            Ninstances = mstForecastDto.NInstancesSpecified ? MyUtilities.ToString(mstForecastDto.NInstances) : null;

            CompCode = mstForecastDto.compCode;

            CompModCode = mstForecastDto.compModCode;

            EquipNo = mstForecastDto.equipNo;

            HideSuppressed = mstForecastDto.hideSuppressedSpecified ? MyUtilities.ToString(mstForecastDto.hideSuppressed) : null;

            MaintSchTask = mstForecastDto.maintSchTask;

            Rec700Type = mstForecastDto.rec700Type;

            ShowRelated = mstForecastDto.showRelatedSpecified ? MyUtilities.ToString(mstForecastDto.showRelated) : null;
        }

        public MstService.MSTForecastDTO ToDto()
        {
            var item = new MstService.MSTForecastDTO();

            item.NInstancesSpecified = Ninstances != null;

            item.NInstances = Convert.ToDecimal(Ninstances);

            item.compCode = !string.IsNullOrWhiteSpace(CompCode) ? CompCode : null;

            item.compModCode = !string.IsNullOrWhiteSpace(CompModCode) ? CompModCode : null;

            item.equipNo = EquipNo;

            item.hideSuppressedSpecified = HideSuppressed != null;

            item.hideSuppressed = MyUtilities.IsTrue(HideSuppressed);

            item.maintSchTask = MaintSchTask;

            item.rec700Type = Rec700Type;

            item.showRelatedSpecified = ShowRelated != null;

            ShowRelated = MyUtilities.ToString(item.showRelated);

            return item;
        }
    }
}
