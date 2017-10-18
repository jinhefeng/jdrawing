using JSolidworks.Func;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace JSolidworks
{
    public class JSolidworksPart
    {
        internal SldWorks swApp;
        public ModelDoc2 swDoc;
        internal string FilePath;
        internal Dictionary<string, object> Sizes = new Dictionary<string, object>();
        private Dictionary<string, object> _Sizes = new Dictionary<string, object>();
        internal Dictionary<string, string> Attrs = new Dictionary<string, string>();
        private Dictionary<string, string> _Attrs = new Dictionary<string, string>();
        private swDocumentTypes_e FileType;
        public JSolidworksPart(string _FilePath)
        {
            this.FilePath = _FilePath;
        }
        public void Load()
        {
            this.swDoc = OpenSWDoc(this.FilePath);
            #region 00. 获取方程式的值
            Sizes.Clear();
            EquationMgr emgr = (EquationMgr)swDoc.GetEquationMgr();
            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                //在这里想办法只获取全局变量
                string s = emgr.get_Equation(i);
                if (s.Contains("@"))
                {
                    continue;
                }
                s = s.Replace("\"", "");
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();
                s = s.ToUpper();
                object v = emgr.get_Value(i);

                if (Sizes.ContainsKey(s))
                {
                    Sizes[s] = v;
                }
                else
                {
                    Sizes.Add(s, v);
                }
            }
            _Sizes = new Dictionary<string, object>(Sizes);
            #endregion
            #region 10. 获取零部件属性的值
            Attrs.Clear();
            object[] PropNames;
            string cfgname2 = swDoc.ConfigurationManager.ActiveConfiguration.Name;
            CustomPropertyManager cpmMdl = swDoc.Extension.get_CustomPropertyManager(cfgname2);
            if(cpmMdl.Count>0)
            {
                PropNames = (object[])cpmMdl.GetNames();
                string outval = "";
                string sovVal = "";
                foreach (object AtrName in PropNames)
                {
                    string s = AtrName.ToString();
                    cpmMdl.Get2(AtrName.ToString(), out outval, out sovVal);
                    if (Attrs.ContainsKey(s))
                    {
                        Attrs[s] = sovVal;
                    }
                    else
                    {
                        Attrs.Add(s, sovVal);
                    }
                }
            }
            _Attrs = new Dictionary<string, string>(Attrs);
            #endregion
        }
        public void Commit()
        {
            #region 00. 提交尺寸
            EquationMgr emgr = (EquationMgr)swDoc.GetEquationMgr();
            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                //在这里想办法只获取全局变量
                string s = emgr.get_Equation(i);
                if (s.Contains("@"))
                {
                    continue;
                }
                s = s.Replace("\"", "");
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();
                s = s.ToUpper();
                if (Sizes.ContainsKey(s))
                {
                    if (Sizes[s] != _Sizes[s])
                    {
                        string Equation = JswFunc.FormatEquation(s, Sizes[s].ToString());
                        emgr.set_Equation(i, Equation);
                    }
                }else
                {
                    if(_Sizes.ContainsKey(s))
                    {
                        emgr.Delete(i);
                        i--;
                        ic--;
                    }
                }
            }
            //foreach (string key in Sizes.Keys)
            //{
            //    if (Sizes[key] != _Sizes[key])
            //    {
            //        JswFunc.SetEquationValueByName(swDoc, key, Sizes[key]);
            //    }
            //}
            _Sizes = new Dictionary<string, object>(Sizes);
            #endregion
            #region 10. 提交属性
            string cfgname2 = swDoc.ConfigurationManager.ActiveConfiguration.Name;
            CustomPropertyManager cpmMdl = swDoc.Extension.get_CustomPropertyManager(cfgname2);
            //修改属性
            foreach (string key in Attrs.Keys)
            {
                if(_Attrs.ContainsKey(key))
                {
                    if (Attrs[key] != _Attrs[key])
                    {
                        cpmMdl.Set(key, Attrs[key]);
                    }
                }else
                {
                    //如果该Key在_Attrs中找不到，表示为新增的属性
                    int outval = 30;
                    string sovVal = "";
                    // cpmMdl.Set(AtrName, Value);
                    cpmMdl.Add2(key, outval, Attrs[key]);
                    
                }
                
            }
            //清理已经删除的属性
            foreach(string key in _Attrs.Keys)
            {
                if (!Attrs.ContainsKey(key))
                {
                    //旧属性不在新列表中，表示该属性已被删除。
                    cpmMdl.Delete(key);
                }
            }
            _Attrs = new Dictionary<string, string>(Attrs);
            #endregion
            
        }
        public void Save()
        {
            swDoc.Save();
        }
        public void Rebuild()
        {
            swDoc.EditRebuild3();
        }
        public void CloseDoc()
        {
            if (this.FilePath != null)
            {
                try
                {
                    this.swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    swApp.CloseDoc(this.FilePath);
                }
                catch
                {
                    //如果不能获取SolidWorks进程，说明软件没有打开，说明文件也没有打开
                }
                
            }
        }
        public object GetSize(string _SizeName)
        {
            if (Sizes.ContainsKey(_SizeName.ToUpper()))
            {
                return Sizes[_SizeName.ToUpper()];
            }
            return null;
        }
        public void SetSize(string _SizeName,object _Size)
        {
            Sizes[_SizeName.ToUpper()] = _Size;
        }
        public string GetAttr(string _AttrName)
        {
            if (Attrs.ContainsKey(_AttrName))
            {
                return Attrs[_AttrName];
            }
            return null;
        }
        public void SetAttr(string _AttrName, string _AttrValue)
        {
            if (Attrs.ContainsKey(_AttrName))
            {
                Attrs[_AttrName] = _AttrValue;
            }else
            {
                Attrs.Add(_AttrName, _AttrValue);
            }
        }
        public void DelAttr(string _AttrName)
        {
            if (Attrs.ContainsKey(_AttrName))
            {
                Attrs.Remove(_AttrName);
            }
        }
        public void ClearSize()
        {
            Sizes.Clear();
            EquationMgr emgr = (EquationMgr)swDoc.GetEquationMgr();
            while (emgr.GetCount() > 0)
            {
                emgr.Delete(0);
            }
        }
        #region 10. 中间过程
        private void OpenDoc(string _FilePath)
        {
            this.FilePath = _FilePath;
            swDoc = OpenSWDoc(_FilePath);
            #region 00. 获取方程式的值
            Sizes.Clear();
            EquationMgr emgr = (EquationMgr)swDoc.GetEquationMgr();
            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                //在这里想办法只获取全局变量
                string s = emgr.get_Equation(i);
                if (s.Contains("@"))
                {
                    continue;
                }
                s = s.Replace("\"", "");
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();
                s = s.ToUpper();
                object v = emgr.get_Value(i);

                if (Sizes.ContainsKey(s))
                {
                    Sizes[s] = v;
                }
                else
                {
                    Sizes.Add(s, v);
                }
            }
            _Sizes = new Dictionary<string, object>(Sizes);
            #endregion
            #region 10. 获取零部件属性的值
            Attrs.Clear();
            _Attrs = new Dictionary<string, string>(Attrs);
            #endregion
            //CloseDoc();
        }
        private ModelDoc2 OpenSWDoc(string _FilePath)
        {
            try
            {
                this.swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                this.swApp = new SldWorks();
            }
            this.FileType = JswFunc.GetswFileType(_FilePath);
            int longstatus = 0;
            int longwarnings = 0;
            return (ModelDoc2)this.swApp.OpenDoc6(_FilePath, (int)this.FileType, 0, "", ref longstatus, ref longwarnings); //打开SW文件
        }
        #endregion
        #region 20. 暂时不用的方法
        private void DelSize(string _SizeName)
        {
            //暂时没有必要,不开放此方法
            string s = _SizeName.ToUpper();
            if (Sizes.ContainsKey(s))
            {
                Sizes.Remove(s);
            }
            //EquationMgr emgr = (EquationMgr)swDoc.GetEquationMgr();
            //int ic = emgr.GetCount();
            //for (int i = 0; i < ic; i++)
            //{
            //    //在这里想办法只获取全局变量
            //    string s = emgr.get_Equation(i);
            //    s = s.Replace("\"", "");
            //    s = s.Remove(s.IndexOf('='));
            //    s = s.Trim();
            //    s = s.ToUpper();
            //    double v = emgr.get_Value(i);
            //    if (s == _SizeName.ToUpper())
            //    {
            //        emgr.Delete(i);
            //        return;
            //    }
            //}
        }
        #endregion
    }
    public class JSolidworksAsemble : JSolidworksPart
    {
        public List<object> Parts = new List<object>();
        public JSolidworksAsemble(string _FilePath) : base(_FilePath)
        {
            //先执行JSolidworksPart中的构造函数
            this.GetParts();
            //
        }
        private void GetParts()
        {
            //获取该文件下的子零部件
            GetChildren(this);
        }
        private void GetChildren(JSolidworksAsemble _JswDoc)
        {
            //this.Parts.Add((JSolidworksAsemble)_JswDoc); //先将此文件加入到子零部件列表中
                                                         //如果是装配体
            if(_JswDoc.swDoc==null)
            {
                _JswDoc.Load();
            }
            Configuration swConf = (Configuration)_JswDoc.swDoc.GetActiveConfiguration();
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
                object[] Children_List = swRootComp.GetChildren();
                foreach (object _doc in Children_List)
                {

                    ModelDoc2 t_doc = ((Component2)_doc).GetModelDoc2();
                    if (t_doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                    {
                        //子件是
                        string _Path=t_doc.GetPathName();
                        JSolidworksPart JPart = new JSolidworksPart(t_doc.GetPathName());
                        JPart.Load();
                        _JswDoc.Parts.Add(JPart);
                    }
                    else
                    {
                        if (t_doc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                        {
                            JSolidworksAsemble JPart = new JSolidworksAsemble(t_doc.GetPathName());
                            JPart.Load();
                            _JswDoc.Parts.Add(JPart);
                        }
                    }
                }
            }
            return;
        }
        //private void GetChildren(object _JswDoc)
        //{
        //    if (_JswDoc is JSolidworksAsemble)
        //    {
        //        this.Parts.Add((JSolidworksAsemble)_JswDoc); //先将此文件加入到子零部件列表中
        //                                                     //如果是装配体
        //        ModelDoc2 _swDoc = ((JSolidworksAsemble)_JswDoc).swDoc;

        //        Configuration swConf = (Configuration)_swDoc.GetActiveConfiguration();
        //        if (swConf == null)
        //        {
        //            return;
        //        }
        //        Component2 swRootComp = (Component2)swConf.GetRootComponent();

        //        if (swRootComp == null)
        //        {
        //            return;
        //        }
        //        else
        //        {
        //            object[] Children_List = swRootComp.GetChildren();
        //            foreach (object _doc in Children_List)
        //            {

        //                ModelDoc2 t_doc = ((Component2)_doc).GetModelDoc2();
        //                if (t_doc.GetType() == (int)swDocumentTypes_e.swDocPART)
        //                {
        //                    //子件是
        //                    JSolidworksPart JPart = new JSolidworksPart(t_doc.GetTitle());
        //                    ((JSolidworksAsemble)_JswDoc).Parts.Add(JPart);
        //                }
        //                else
        //                {
        //                    if (t_doc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
        //                    {
        //                        //子件是
        //                        JSolidworksAsemble JPart = new JSolidworksAsemble(t_doc.GetTitle());
        //                        JPart.GetChildren(JPart);
        //                        ((JSolidworksAsemble)_JswDoc).Parts.Add(JPart);
        //                    }
        //                }
        //            }
        //        }

        //        (_JswDoc as JSolidworksAsemble).GetChildren(_JswDoc);
        //    }
        //    else
        //    {
        //        if (_JswDoc is JSolidworksPart)
        //        {
        //            //如果是零件

        //            this.Parts.Add(_JswDoc);
        //        }
        //    }
        //    return;
        //}
    }
    public class JSolidworksFunc
    {
        public static void SetSizeById(string id,string SizeName,object SizeValue,object swDoc,bool isDeep)
        {
            //指定一个零件的ID和母文件，即可对其的尺寸进行修改。
            switch(swDoc.GetType().ToString())
            {
                case "JSolidworks.JSolidworksPart":
                    JSolidworksPart JPart = swDoc as JSolidworksPart;
                    if (JPart.GetAttr("id")==id)
                    {
                        JPart.SetSize(SizeName, SizeValue);
                    }
                    break;
                case "JSolidworks.JSolidworksAsemble":
                    JSolidworksAsemble JAsemble = swDoc as JSolidworksAsemble;
                    if (JAsemble.GetAttr("id") == id)
                    {
                        JAsemble.SetSize(SizeName, SizeValue);
                    }else
                    {
                        if (isDeep)
                        {
                            if (JAsemble.Parts.Count > 0)
                            {
                                foreach (object _swDoc in JAsemble.Parts)
                                {
                                    SetSizeById(id, SizeName, SizeValue, _swDoc, isDeep);
                                }
                            }
                        }
                    }
                    break;
                default:
                    break;

            }
        }
        public static void SetAttrById(string id, string AttrName, string AttrValue, object swDoc, bool isDeep)
        {
            //指定一个零件的ID和母文件，即可对其的属性进行修改。
            switch (swDoc.GetType().ToString())
            {
                case "JSolidworks.JSolidworksPart":
                    JSolidworksPart JPart = swDoc as JSolidworksPart;
                    if (JPart.GetAttr("id") == id)
                    {
                        JPart.SetAttr(AttrName, AttrValue);
                    }
                    break;
                case "JSolidworks.JSolidworksAsemble":
                    JSolidworksAsemble JAsemble = swDoc as JSolidworksAsemble;
                    if (JAsemble.GetAttr("id") == id)
                    {
                        JAsemble.SetAttr(AttrName, AttrValue);
                    }
                    else
                    {
                        if (isDeep)
                        {
                            if (JAsemble.Parts.Count > 0)
                            {
                                foreach (object _swDoc in JAsemble.Parts)
                                {
                                    SetAttrById(id, AttrName, AttrValue, _swDoc, isDeep);
                                }
                            }
                        }
                    }
                    break;
                default:
                    break;

            }
        }
    }
}
