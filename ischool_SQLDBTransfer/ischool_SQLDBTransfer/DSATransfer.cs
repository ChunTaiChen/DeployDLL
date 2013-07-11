using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using FISCA.DSAClient;
using System.Windows.Forms;
using System.IO;


namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 與 DSA 交換
    /// </summary>
    public class DSATransfer
    {
        string _AccessPoint = "";
        string _ContractName = "";
        string _UserName = "";
        string _Password = "";
        string _ServiceName = "";
        XElement _Req;

        public DSATransfer(string AccessPoint, string ContractName, string UserName, string Password, string ServiceName, XElement Request)
        {
            _AccessPoint = AccessPoint;
            _ContractName = ContractName;
            _UserName = UserName;
            _Password = Password;
            _ServiceName = ServiceName;
            _Req = Request;
        }


        /// <summary>
        /// 取得回傳XML資料
        /// </summary>
        /// <returns></returns>
        public XElement GetResponse()
        {
            XElement elm = null;
            Envelope rsp = null;

            Console.WriteLine("test");
            //try
            //{
            Connection cn = new Connection();
            cn.Connect(_AccessPoint, _ContractName, _UserName, _Password);

            XmlHelper req = new XmlHelper(_Req.ToString());

            rsp = cn.SendRequest(_ServiceName, new Envelope(req));

            if (rsp.Body.XmlString != "")
            {
                XmlHelper xmlrsp = new XmlHelper(rsp.Body);
                elm = XElement.Parse(xmlrsp.XmlString);
            }
            //}
            //catch (DSAServerException ex)
            //{
            //    string fileName = "ResponseError" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".txt";
            //    string file_path = Application.StartupPath + "\\" + fileName;
            //    StreamWriter sw = File.CreateText(file_path);
            //    sw.Write(FISCA.ErrorReport.Generate(ex));
            //    sw.Close();
            //    throw;
            //}

            //catch (Exception ex)
            //{
            //    string fileName = "ResponseError" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".txt";
            //    string file_path = Application.StartupPath + "\\" + fileName;
            //    StreamWriter sw = File.CreateText(file_path);
            //    sw.Write(FISCA.ErrorReport.Generate(ex));
            //    sw.Close();

            //    Global._ExceptionManager.AddMessage(_AccessPoint + "," + _ContractName + "," + _ServiceName + "," + ex.Message + "::" + FISCA.ErrorReport.Generate(ex));
            //    Global._ExceptionManager.AddMessage(_AccessPoint + "," + _ContractName + "," + _ServiceName + "," + ex.Message + rsp.XmlString);                
            //    //throw new D{;AServerException();
            //    throw;
            //}

            //if (elm == null)
            //{
            //    //Global._ExceptionManager.AddMessage(_AccessPoint + "," + _ContractName + "," + _ServiceName+"," + rsp.XmlString);
            //}
            return elm;
        }

    }
}
