using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 資料格式轉換
    /// </summary>
    public class DataTransfer
    {
        /// <summary>
        /// 待轉換資料
        /// </summary>
        XElement _XmlData;
        /// <summary>
        /// 設定檔
        /// </summary>
        XElement _XmlConfig;
        /// <summary>
        /// 儲存轉換後的SQL
        /// </summary>
        DataStore _DataStore;

        List<string> _FieldList;
        List<string> _DataList;

        string _DBName = "";

        public DataTransfer()
        {
            _DataStore = new DataStore();
            _FieldList = new List<string>();
            _DataList = new List<string>();
        }

        /// <summary>
        /// 傳入轉換資料
        /// </summary>
        /// <param name="config"></param>
        /// <param name="Data"></param>
        public void SetData(XElement config, XElement Data,string DBName)
        {
            _XmlConfig = config;
            _XmlData = Data;
            _DBName = DBName;
            ConvertData();
        }

        /// <summary>
        /// 轉換資料
        /// </summary>
        private void ConvertData()
        {
            string tableName = _XmlConfig.Element("DBTableName").Value;
            string RspElementName = _XmlConfig.Element("RspElementName").Value;
            string Action = _XmlConfig.Element("Action").Value;

                if (Action.ToLower() == "update")
                {
                    XElement ex = new XElement("ex"); 
                    XElement ex1 = new XElement("ex1");
                    string err = "";
                    //_DataStore.AddData(tableName, "use " + _DBName);
                    try
                    {
                       
                        foreach (XElement elmData in _XmlData.Elements(RspElementName))
                        {
                            foreach (XElement elm in _XmlConfig.Element("FieldMapping").Elements("Field"))
                            {
                                ex = elm;
                                ex1 = elmData;

                                if (elmData.Element(elm.Attribute("Source").Value) != null)
                                {
                                    if (elmData.Element(elm.Attribute("Source").Value).Value == elm.Attribute("SourceField").Value)
                                    {
                                        string query = " update " + tableName + " set ";
                                        if(elmData.Element(elm.Attribute("SourceValue").Value) ==null )
                                            query += elm.Attribute("Target").Value + "=N'' where " + elm.Attribute("TargetConditionField").Value + "=N'" + elmData.Element(elm.Attribute("SourceConditionField").Value).Value + "'";
                                        else
                                            query += elm.Attribute("Target").Value + "=N'" + elmData.Element(elm.Attribute("SourceValue").Value).Value + "' where " + elm.Attribute("TargetConditionField").Value + "=N'" + elmData.Element(elm.Attribute("SourceConditionField").Value).Value + "'";
                                        _DataStore.AddData(tableName + "_update", query);
                                        err = query;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception Exception)
                    {
                        Global._ExceptionManager.AddMessage("update::" + ex.ToString()+","+ex1.ToString ()+","+err);
                    }                        
                }
                else
                {
                    // insert
                    //_DataStore.AddData(tableName, "use " + _DBName);                



                    Dictionary<string, string> FieldMapDict = new Dictionary<string, string>();
                    Dictionary<string, string> FieldTypeDict = new Dictionary<string, string>();
                    // 設定檔
                    foreach (XElement elm in _XmlConfig.Element("FieldMapping").Elements("Field"))
                    {
                        string key = elm.Attribute("Source").Value;

                        if (!FieldMapDict.ContainsKey(key))
                        {
                            FieldMapDict.Add(key, elm.Attribute("Target").Value);
                            FieldTypeDict.Add(key, elm.Attribute("Type").Value);
                        }
                    }
                    // 資料
                    foreach (XElement elmData in _XmlData.Elements(RspElementName))
                    {
                        _FieldList.Clear();
                        _DataList.Clear();

                        foreach (string key in FieldMapDict.Keys)
                        {
                            _FieldList.Add(FieldMapDict[key]);
                            if (elmData.Element(key) != null)
                            {
                                if (FieldTypeDict[key].ToLower() == "integer")
                                {
                                    if (string.IsNullOrEmpty(elmData.Element(key).Value))
                                        _DataList.Add("null");
                                    else
                                        _DataList.Add(elmData.Element(key).Value);
                                }
                                else
                                    _DataList.Add("N'" + elmData.Element(key).Value + "'");
                            }
                            else
                            {
                                // 存入系統存取點
                                if (key == "SYS_AccessName")
                                    _DataList.Add("N'" + Global._DSACurrectName + "'");
                                else
                                    _DataList.Add("");
                            }
                        }
                        // 主要單獨處理教師帳號有重複問題
                        string AddStr = "if not exists(select login_name from teacher where login_name=" + _DataList[2] + ") ";
                        if (tableName == "teacher")
                            _DataStore.AddData(tableName, AddStr + "insert into " + tableName + "(" + string.Join(",", _FieldList.ToArray()) + ") values(" + string.Join(",", _DataList.ToArray()) + ");");
                        else
                            _DataStore.AddData(tableName, "insert into " + tableName + "(" + string.Join(",", _FieldList.ToArray()) + ") values(" + string.Join(",", _DataList.ToArray()) + ");");
                    }
                }
          

        }

        /// <summary>
        /// 取得儲存資料
        /// </summary>
        /// <returns></returns>
        public DataStore GetData()
        {
          
            return _DataStore;
        }

    }
}
