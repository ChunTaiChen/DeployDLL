using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ischool_SQLDBTransfer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConfigManager cm = new ConfigManager();
            foreach(XElement elmAccessPoint in cm.GetAccessPointList())
            foreach (XElement elmJob in cm.GetTransferJobList())
            {
                DSATransfer dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, elmJob.Element("ServiceName").Value, elmJob.Element("Request"));
                DataTransfer dt = new DataTransfer();
                dt.SetData(elmJob, dsat.GetResponse(), elmAccessPoint.Attribute("DBName").Value);
                DBTransfer dbt = new DBTransfer(elmAccessPoint.Attribute("DBServerName").Value, elmAccessPoint.Attribute("DBName").Value, elmAccessPoint.Attribute("DBUserName").Value, elmAccessPoint.Attribute("DBPassword").Value);
                dbt.ExecuteQuerySQLByDataStore(dt.GetData());
            }
            


            //DSATransfer dt = new DSATransfer("http://test.iteacher.tw/cs4/test_sh_d", "ischool.transfer01", "test", "test", "data.GetTeacherList");
            //DBTransfer dbt = new DBTransfer("ct2\\sqlexpress", "testdb01", "test", "test");
            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine("use testdb01;");            
            //sb.AppendLine("insert into class(id,class_name) values(1,'1')");
            
        }
    }
}
