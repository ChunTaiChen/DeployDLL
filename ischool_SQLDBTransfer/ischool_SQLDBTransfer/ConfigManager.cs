using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 與設定檔交換
    /// </summary>
    public class ConfigManager
    {
        XElement _ConfigXml;

        /// <summary>
        /// 初始化解析
        /// </summary>
        public ConfigManager()
        {
            try
            {
                _ConfigXml = XElement.Load(Application.StartupPath + "\\Configuration.xml");
                if (_ConfigXml == null)
                {
                    Global._ExceptionManager.AddMessage("沒有設定資料");
                    return;
                }
            }
            catch (Exception ex)
            {
                Global._ExceptionManager.AddMessage("設定檔解析失敗!" + ex.Message);
            }
        }

        /// <summary>
        /// 取得服務存取點與資料庫資料
        /// </summary>
        /// <returns></returns>
        public List<XElement> GetAccessPointList()
        {
            List<XElement> retVal = new List<XElement>();
            retVal =(from data in _ConfigXml.Element("DSADBLocation").Elements("AP") select data ).ToList();
            return retVal;

        }

        /// <summary>
        /// 取得交換工作內容
        /// </summary>
        /// <returns></returns>
        public List<XElement> GetTransferJobList()
        {
            List<XElement> retVal = new List<XElement>();
            retVal = (from data in _ConfigXml.Elements("TransferJob") select data).ToList();
            return retVal;
        }

        /// <summary>
        /// 取得所有XML
        /// </summary>
        /// <returns></returns>
        public XElement GetAll()
        {
            return _ConfigXml;
        }
    }
}
