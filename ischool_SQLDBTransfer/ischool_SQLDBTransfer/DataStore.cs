using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 資料儲存
    /// </summary>
    public class DataStore
    {
        /// <summary>
        /// 儲存資料  < tableName , List<SQL string> >
        /// </summary>
        Dictionary<string, List<string>> _DataStoreDict;

        public DataStore()
        {
            _DataStoreDict = new Dictionary<string, List<string>>();
        }

        /// <summary>
        /// 新增資料
        /// </summary>
        /// <param name="name"></param>
        /// <param name="Value"></param>
        public void AddData(string name, string Value)
        {
            if (!_DataStoreDict.ContainsKey(name))
                _DataStoreDict.Add(name,new List<string>());                

            _DataStoreDict[name].Add(Value);
        }

        /// <summary>
        /// 透過名稱取得List資料內容
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public List<string> GetDataListByName(string name)
        {
            if (_DataStoreDict.ContainsKey(name))
                return _DataStoreDict[name];
            else
                return new List<string>();
        }

        /// <summary>
        /// 取得所有資料
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, List<string>> GetAllDataDict()
        {
            return _DataStoreDict;
        }

        /// <summary>
        /// 依名稱取得資料筆數
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public int GetDataListCountByName(string name)
        {
            if (_DataStoreDict.ContainsKey(name))
                return _DataStoreDict[name].Count;
            else
                return 0;        
        }

        public void DictinctData(string name)
        {
            if (_DataStoreDict.ContainsKey(name))
            {
                List<string> temp = _DataStoreDict[name].Distinct().ToList();
                _DataStoreDict[name] = temp;
            }        
        
        }
    }
}
