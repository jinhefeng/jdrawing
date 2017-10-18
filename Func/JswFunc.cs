using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSolidworks.Func
{
    public class JswFunc
    {
        #region 00. 方程式
        #region 获取全局变量
        /// <summary>
        /// 获取方程式中全局变量的值，Golbal Variables
        /// </summary>
        /// <param name="VariableName">全局变量的名称</param>
        /// <param name="Part">模型</param>
        /// <returns>如果没有找到这个全局变量，返回-1</returns>
        public double getGolbalVariablesValue(string VariableName, ModelDoc2 Part)
        {
            EquationMgr emgr = (EquationMgr)Part.GetEquationMgr();
            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                string s = emgr.get_Equation(i);
                s = s.Replace("\"", "");
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();
                if (s.ToLower() == VariableName.ToLower())
                {
                    return emgr.get_Value(i);
                }
            }
            return -1;
        }
        #endregion
        #region 修改全局变量
        public static bool setGolbalVariablesValue(string VariableName, ModelDoc2 Part, string Value)
        {
            EquationMgr emgr = (EquationMgr)Part.GetEquationMgr();
            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                string s = emgr.get_Equation(i);
                s = s.Replace("\"", "");
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();
                if (s.ToLower() == VariableName.ToLower())
                {
                    emgr.set_Equation(i, Value);
                    return true;
                }
            }
            return false;
        }
        public static string FormatEquation(string Key, string Value)
        {
            string returnValue = null;
            returnValue = "\"" + Key + "\"=\"" + Value + "\"";
            return returnValue;
        }
        #endregion
        #region 01.读取方程式的值
        public static void GetEquationValueByName(ModelDoc2 Part,string VariableName,ref object Value)
        {
            /// <summary>
            /// 获取方程式中全局变量的值，Golbal Variables
            /// </summary>
            /// <param name="VariableName">全局变量的名称</param>
            /// <param name="Part">模型</param>
            /// <returns>如果没有找到这个全局变量，返回-1</returns>

            EquationMgr emgr = (EquationMgr)Part.GetEquationMgr();
            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                string s = emgr.get_Equation(i);
                s = s.Replace("\"", "");
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();
                if (s.ToLower() == VariableName.ToLower())
                {
                    Value=emgr.get_Value(i);
                    return;
                }
            }
            Value = -1;
        }
        #endregion
        #region 02.写入方程式的值
        public static void SetEquationValueByName(ModelDoc2 Part, string VariableName, object Value)
        {
            string Equation = FormatEquation(VariableName, Value.ToString());
            setGolbalVariablesValue(VariableName, Part, Equation);
        }
        #endregion
        #endregion
        #region 10. 零部件属性
        #region 根据ModelDoc得到零件的属性
        public static string GetAttValueByName(string AtrName, ModelDoc2 swModel)
        {
            string cfgname2 = swModel.ConfigurationManager.ActiveConfiguration.Name;

            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgname2);
            string outval = "";
            string sovVal = "";
            cpmMdl.Get2(AtrName, out outval, out sovVal);
            return sovVal;
        }
        #endregion
        #region 写入零件属性
        public static string SetAttValueByName(string AtrName, ModelDoc2 swModel, string Value)
        {
            string cfgname2 = swModel.ConfigurationManager.ActiveConfiguration.Name;

            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgname2);
            //string outval = "";
            string sovVal = "";
            cpmMdl.Set(AtrName, Value);

            return sovVal;
        }
        #endregion
        #region 增加零件属性
        public static string AddAttValueByName(string AtrName, ModelDoc2 swModel, string Value)
        {
            string cfgname2 = swModel.ConfigurationManager.ActiveConfiguration.Name;

            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgname2);
            int outval = 30;
            string sovVal = "";
            // cpmMdl.Set(AtrName, Value);
            cpmMdl.Add2(AtrName, outval, Value);
            return sovVal;

        }
        #endregion
        #region 删除零件属性
        public static string DelAttValueByName(string AtrName, ModelDoc2 swModel, string Value)
        {
            string cfgname2 = swModel.ConfigurationManager.ActiveConfiguration.Name;

            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgname2);
            int outval = 30;
            string sovVal = "";
            // cpmMdl.Set(AtrName, Value);
            cpmMdl.Delete(AtrName);
            return sovVal;

        }
        #endregion
        #endregion
        #region 20. 零部件关系
        #region 获取特征树中的零件
        public static void GetPartsFromAssembly(ModelDoc2 doc, ref List<ModelDoc2> Parts_List)
        {
            if (doc == null)
            {
                return;
            }

            int docType = doc.GetType();
            if (docType == (int)swDocumentTypes_e.swDocPART)
            {
                Parts_List.Add(doc);
                return;
            }
            //object haha = doc.get;

            Configuration swConf = (Configuration)doc.GetActiveConfiguration();
            if (swConf == null)
            {
                return;
            }
            Component2 swRootComp = (Component2)swConf.GetRootComponent();

            if (swRootComp == null)
            {
                return;
            }
            else
            {
                Parts_List.Add(doc);
                object[] haha = swRootComp.GetChildren();
                foreach (object hh in haha)
                {
                    Component2 hh2 = (Component2)hh;
                    ModelDoc2 hhhh = hh2.GetModelDoc2();
                    GetPartsFromAssembly(hhhh, ref Parts_List);
                }
            }
        }
        #endregion
        #region 获取装配体中的装配关系
        public void GetParametersFromAssembly(ModelDoc2 doc, ref List<ModelDoc2> Parameters_List)
        {
            if (doc == null)
            {
                return;
            }

            int docType = doc.GetType();
            if (docType == (int)swDocumentTypes_e.swDocPART)
            {
                Parameters_List.Add(doc);
                return;
            }
            //object haha = doc.get;

            Configuration swConf = (Configuration)doc.GetActiveConfiguration();
            if (swConf == null)
            {
                return;
            }
            Component2 swRootComp = (Component2)swConf.GetRootComponent();
            int swRootComp1 = (int)swConf.GetParameterCount();
            if (swRootComp1 == null)
            {
                return;
            }
            else
            {
                Parameters_List.Add(doc);
                object[] haha = swRootComp.GetChildren();

                foreach (object hh in haha)
                {
                    Component2 hh2 = (Component2)hh;
                    ModelDoc2 hhhh = hh2.GetModelDoc2();
                    GetParametersFromAssembly(hhhh, ref Parameters_List);
                }
            }
        }
        #endregion
        #endregion
        #region 30. 通用
        public static swDocumentTypes_e GetswFileType(string _FilePath)
        {
            swDocumentTypes_e FileType;
            switch (_FilePath.Split('.').Last().ToUpper())
            {
                case "SLDPRT":
                    FileType = swDocumentTypes_e.swDocPART;
                    break;
                case "SLDASM":
                    FileType = swDocumentTypes_e.swDocASSEMBLY;
                    break;
                case "SLDDRW":
                    FileType = swDocumentTypes_e.swDocDRAWING;
                    break;
                default:
                    FileType = swDocumentTypes_e.swDocNONE;
                    break;
            }
            return FileType;
        }
        #endregion
        #region 99. 其他相关
        #region //读取TextBox中的组合值
        public double[] Range(string JianJu)
        {
            //将,转换成空格 
            JianJu = JianJu.Replace(',', ' ');
            string[] Separation = JianJu.Split(' ');//为每个数据分配数组资源 
            double[] returnValue = new double[20];
            int j = 0;
            foreach (string every in Separation)
            {
                try
                {
                    returnValue[j] = Convert.ToDouble(every);
                    j++;
                }
                catch
                {
                    try
                    {
                        string[] Count = every.Split('*');
                        for (int k = 0; k < Convert.ToInt32(Count[0]); k++)
                        {
                            returnValue[j] = Convert.ToDouble(Count[1]);
                            j++;
                        }
                    }
                    catch { }
                }
            }
            return returnValue;
        }
        #endregion
        #endregion
    }
}
