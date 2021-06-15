using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SS = EllipseSubAssetGeneralInfoExcelAddIn.ScreenService;

namespace EllipseSubAssetGeneralInfoExcelAddIn
{
    public class Ellipse
    {
        public static String ERR_FLOW_SCREEN = "ERROR EN FLUJO DE PANTALLAS.";
        public static String OK = "1";
        public static String CANCEL = "3";

        public static String TRANSMIT = "1";
        public static String F1_KEY = "2";
        public static String F2_KEY = "3";
        public static String F3_KEY = "4";
        public static String F4_KEY = "5";
        public static String F5_KEY = "6";
        public static String F6_KEY = "7";
        public static String F7_KEY = "8";
        public static String F8_KEY = "9";
        public static String F9_KEY = "10";
        public static String F10_KEY = "11";

        public Ellipse()
        {
        }

        public static void destroyMSO(SS.ScreenService proxy, SS.OperationContext op, SS.ScreenSubmitRequestDTO ss)
        {
            if (proxy != null)
            {
                if (op != null)
                {
                    if (ss != null)
                    {
                        ss.screenFields = null;
                        ss.screenKey = Ellipse.CANCEL;
                        proxy.submit(op, ss);
                    }
                }
            }
        }

        public static void destroyMSO(SS.ScreenService proxy, SS.OperationContext op, SS.ScreenSubmitRequestDTO ss, SS.ScreenDTO mso)
        {
            if (mso != null && mso.mapName != null && !mso.mapName.Equals(""))
            {
                if (proxy != null)
                {
                    if (op != null)
                    {
                        if (ss != null)
                        {
                            ss.screenFields = null;
                            ss.screenKey = Ellipse.CANCEL;
                            proxy.submit(op, ss);
                        }
                    }
                }
            }
        }

        public static SS.ScreenDTO executeMSO(SS.ScreenService proxy, SS.OperationContext op, SS.ScreenSubmitRequestDTO ss, SS.ScreenNameValueDTO[] screenFields, String option)
        {
            ss.screenFields = screenFields;
            ss.screenKey = option;
            return proxy.submit(op, ss);
        }

        public static SS.ScreenDTO executeScreen(SS.ScreenService proxy, SS.OperationContext op, SS.ScreenSubmitRequestDTO ss, String screen, String screenApp)
        {
            SS.ScreenDTO mso = proxy.executeScreen(op, screen);
            ss = new SS.ScreenSubmitRequestDTO();
            if (!mso.mapName.Equals(screenApp))
            {
                mso = Ellipse.executeMSO(proxy, op, ss, null, Ellipse.CANCEL);
                mso = proxy.executeScreen(op, screen);
            }
            return mso;
        }

        public static SS.ScreenNameValueDTO setMSOFieldValue(String fieldName, String value)
        {
            SS.ScreenNameValueDTO obj = new SS.ScreenNameValueDTO();
            obj.fieldName = fieldName;
            obj.value = value;
            return obj;
        }

        public static String getMSOFieldValue(SS.ScreenFieldDTO[] list, String fieldName)
        {
            String value = "";
            foreach (SS.ScreenFieldDTO obj in list)
            {
                if (obj.fieldName.Trim().Equals(fieldName.Trim()))
                {
                    value = obj.value.Trim();
                    break;
                }
            }
            return value;
        }

        public static String getMSOFieldValue(SS.ScreenDTO mso, String fieldName)
        {
            String value = "";
            SS.ScreenFieldDTO[] list = mso.screenFields;
            foreach (SS.ScreenFieldDTO obj in list)
            {
                if (obj.fieldName.Trim().Equals(fieldName.Trim()))
                {
                    value = obj.value.Trim();
                    break;
                }
            }
            return value;
        }

        public static String getFullMessageError(SS.ScreenDTO mso)
        {
            return "<" + mso.mapName + "><" + mso.currentCursorFieldName + "> " + mso.message;
        }

        public static bool isNotEmptyOrNull(String value)
        {
            if (value != null && !value.Trim().Equals(""))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool isEmptyOrNull(String value)
        {
            return !isNotEmptyOrNull(value);
        }
    }
}