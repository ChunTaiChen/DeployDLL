using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;

namespace ischool_SQLDBTransfer
{
    public class ExceptionManager
    {
        List<string> _MessageList;
        
        public ExceptionManager()
        {
            _MessageList = new List<string>();
        }

        public void AddMessage(string msg)
        {
            msg = DateTime.Now.ToString() + ":" + msg;
            _MessageList.Add(msg);
        }

        public void Clear()
        {
            if (_MessageList == null)
                _MessageList = new List<string>();

            _MessageList.Clear();
        }

        public void Save()
        {
            XElement elmRoot = new XElement("ErrorMessage");
            foreach (string str in _MessageList)
            {
                XElement elm = new XElement("Message");
                elm.SetAttributeValue("Message", str);
                elmRoot.Add(elm);
            }

            string fileName="Exception"+DateTime.Now.Year+"-"+DateTime.Now.Month+"-"+DateTime.Now.Day+".xml";

            if (_MessageList.Count > 0)
                elmRoot.Save(Application.StartupPath + "\\" + fileName);
        }

        public int ErrorCount()
        {
            return _MessageList.Count;
        }
    }
}
