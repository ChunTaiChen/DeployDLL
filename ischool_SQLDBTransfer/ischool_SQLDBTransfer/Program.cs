using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ischool_SQLDBTransfer
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            Global._ExceptionManager.Clear();
            ProcessManager pm = new ProcessManager();
            try
            {
                pm.Start();
            }
            catch(Exception ex)
            {
               // MessageBox.Show(ex.Message);
                Global._ExceptionManager.AddMessage(ex.Message);                
            }
                       
            Global._ExceptionManager.Save();
        }
    }
}
