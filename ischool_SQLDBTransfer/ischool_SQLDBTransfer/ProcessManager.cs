using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows.Forms;
using System.Data;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 程序管理者
    /// </summary>
    public class ProcessManager
    {
        ConfigManager _ConfMang ;
        DBTransfer _dbt;
        DSATransfer _dsat;
        DataTransfer _dt;
        StringBuilder sb_SQL;
        List<XElement>_ElmAccessPoint;
        DateTime _beginDateTime;
        DateTime _endDateTime;
        Stopwatch sp;

        int _IsProcessPass=0;
        int _IsProcessPass2 = 0;
        public ProcessManager()
        {
            _ConfMang = new ConfigManager ();
            sb_SQL= new StringBuilder ();            
            _ElmAccessPoint =_ConfMang.GetAccessPointList();
        }

        /// <summary>
        /// 開始執行
        /// </summary>
        public void Start()
        {
            sp = new Stopwatch();
            _beginDateTime = DateTime.Now;
            _IsProcessPass = 0;
            bool runTruncateTable = true;

            sp.Reset();
            sp.Start();

            // 當第一步已完成開始回傳資料到 ischool
            if (CheckDBStep1())
            {
                // 回傳缺曠相關
                InsertOrUpdateAttendanceToIschool();

                // 回傳成績相關
                InsertOrUpdateScoreToIschool();
                _IsProcessPass2 = 1;
                // 寫入第二步
                WriteDBStep2();

            }
            else
            {
                WriteNoRunStep1();
            }

            if (CheckDBStep3())
            {

                _dt = new DataTransfer();
                string str = "";
                try
                {
                    
                    foreach (XElement elmAccessPoint in _ConfMang.GetAccessPointList())
                    {
                        Global._DSACurrectName = elmAccessPoint.Attribute("Name").Value;

                        XElement Req = new XElement("Request");
                        // 取得預設學年度學期
                        _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, "data.GetDefaultSchoolYearSemester", Req);

                        XElement rsp = _dsat.GetResponse();
                        if (rsp != null)
                        {
                            Global._DefaultSchoolYear = rsp.Element("SystemConfig").Element("DefaultSchoolYear").Value;
                            Global._DefaultSemester = rsp.Element("SystemConfig").Element("DefaultSemester").Value;
                        }

                        foreach (XElement elmJob in _ConfMang.GetTransferJobList())
                        {
                            // 檢查是否可以執行
                            bool CanRunService = false;
                            foreach (XElement el in elmJob.Element("APName").Elements("Name"))
                            {
                                
                                if (el.Value == elmAccessPoint.Attribute("Name").Value)
                                    CanRunService = true;
                            }
                            if (CanRunService)
                            {
                                // 檢查 Req 是否加入預設學年度學期
                                if (elmJob.Element("RequestAddDefault") != null)
                                    if (elmJob.Element("RequestAddDefault").Value.ToLower() == "true")
                                    {
                                        elmJob.Element("Request").SetElementValue("Condition", "");
                                        elmJob.Element("Request").Element("Condition").SetElementValue("SchoolYear", Global._DefaultSchoolYear);
                                        elmJob.Element("Request").Element("Condition").SetElementValue("Semester", Global._DefaultSemester);
                                    }

                                str = elmJob.Element("ServiceName").Value;
                                // 呼叫 Service 取得資料
                                _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, elmJob.Element("ServiceName").Value, elmJob.Element("Request"));
                                // 資料進行轉換
                                _dt.SetData(elmJob, _dsat.GetResponse(), elmAccessPoint.Attribute("DBName").Value);
                            }
                        }
                        // 寫入資料庫 
                        _dbt = new DBTransfer(elmAccessPoint.Attribute("DBServerName").Value, elmAccessPoint.Attribute("DBName").Value, elmAccessPoint.Attribute("DBUserName").Value, elmAccessPoint.Attribute("DBPassword").Value);

                        // 清空 table 內容
                        if (runTruncateTable)
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine("use " + elmAccessPoint.Attribute("DBName").Value);
                            sb.AppendLine("truncate table attendance_ischool");
                            sb.AppendLine("truncate table class");
                            sb.AppendLine("truncate table course");
                            sb.AppendLine("truncate table course_exam");
                            sb.AppendLine("truncate table student");
                            sb.AppendLine("truncate table teacher");
                            sb.AppendLine("truncate table sc_attend");
                            sb.AppendLine("truncate table sc_attend_ischool");
                            _dbt.ExecuteNonQuerySQL(sb);

                        }

                        if (_dbt.ExecuteQuerySQLByDataStore(_dt.GetData()))
                            _IsProcessPass = 1;
                        else
                            _IsProcessPass = 0;
                    }
                }
                catch (Exception ex)
                {
                    Global._ExceptionManager.AddMessage(Global._DSACurrectName + ","+str + ex.Message);
                }

                    
                

                // 新增或更新學期成績
                //InsertOrUpdateSC_Attend();

                // 新增或更新學期成績
                _dbt.ExecuteSpc_SC_Attend();

                sp.Stop();                
                WriteDBStep3();                
            }

            
        }


        /// <summary>
        /// 檢查資料庫系統內當天Step 1是否已執行
        /// </summary>
        /// <returns></returns>
        public bool CheckDBStep1()
        {
            bool pass = false;
            sb_SQL.Clear();
            // 取得設定檔
            if (_ElmAccessPoint.Count > 0)
            {
                _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use " + dbName);
                sb_SQL.AppendLine("select id from transfer_summary where transfer_step=1 and convert(varchar,begin_time,1)=convert(varchar,getdate(),1) and is_successful=1");

                List<string> resultList = _dbt.ExecuteQuerySQL(sb_SQL);
                if (resultList.Count > 0)
                    pass = true;
            }
            return pass;
        }



        /// <summary>
        /// 檢查資料庫系統內當天Step 2是否已執行
        /// </summary>
        /// <returns></returns>
        public bool CheckDBStep2()
        {
            bool pass = false;
            sb_SQL.Clear();
            // 取得設定檔
            if (_ElmAccessPoint.Count > 0)
            {
                _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use "+dbName);
                sb_SQL.AppendLine("select id from transfer_summary where transfer_step=2 and convert(varchar,begin_time,1)=convert(varchar,getdate(),1) and is_successful=1");

                List<string> resultList = _dbt.ExecuteQuerySQL(sb_SQL);
                if (resultList.Count > 0)
                    pass = true;
            }
            return pass;
        }


        /// <summary>
        /// 檢查資料庫系統內當天Step 2是否需要執行
        /// </summary>
        /// <returns></returns>
        public bool CheckDBStep2Do()
        {
            bool pass = true;
            sb_SQL.Clear();
            // 取得設定檔
            if (_ElmAccessPoint.Count > 0)
            {
                _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use " + dbName);
                sb_SQL.AppendLine("select id from transfer_summary where transfer_step=2 and convert(varchar,begin_time,1)=convert(varchar,getdate(),1) and is_successful=1");

                List<string> resultList = _dbt.ExecuteQuerySQL(sb_SQL);
                // 已有執行過
                if (resultList.Count > 0)
                    pass = false;
            }
            return pass;
        }

        /// <summary>
        /// 檢查資料庫系統內當天Step 3是否需要執行
        /// </summary>
        /// <returns></returns>
        public bool CheckDBStep3()
        {            
            bool pass = true;
            sb_SQL.Clear();
            // 取得設定檔
            if (_ElmAccessPoint.Count > 0)
            {
                _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use " + dbName);
                sb_SQL.AppendLine("select id from transfer_summary where transfer_step=3 and convert(varchar,begin_time,1)=convert(varchar,getdate(),1) and is_successful=1");

                List<string> resultList = _dbt.ExecuteQuerySQL(sb_SQL);
                if (resultList == null)
                    resultList = new List<string>();
                // 已有執行過
                if (resultList.Count > 0)
                    pass = false;
            }
            return pass;
        }

        public void WriteDBStep2()
        {
            _endDateTime = DateTime.Now;
            sb_SQL.Clear();
            if (_ElmAccessPoint.Count > 0)
            {
                if (_dbt == null)
                    _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                StringBuilder msg = new StringBuilder();
                msg.Append("'");
                foreach (string str in _dbt.GetData().GetAllDataDict().Keys)
                {
                    foreach (string val in _dbt.GetData().GetAllDataDict()[str])
                    {
                        msg.AppendLine(val);
                    }
                }
                msg.AppendLine("總共費時:" + sp.Elapsed.Minutes + "分" + sp.Elapsed.Seconds + "秒");
                msg.Append("'");
                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use " + dbName);
                sb_SQL.AppendLine("insert into transfer_summary(begin_time,end_time,transfer_step,is_successful,summary) values('" + _beginDateTime.ToString("yyyy/MM/dd HH:mm:ss") + "','" + _endDateTime.ToString("yyyy/MM/dd HH:mm:ss") + "',2," + _IsProcessPass2 + "," + msg.ToString() + ")");


                _dbt.ExecuteNonQuerySQL(sb_SQL);

            }
        }

        /// <summary>
        /// 寫入沒有執行第1步
        /// </summary>
        public void WriteNoRunStep1()
        {
            _endDateTime = DateTime.Now;
            if (_ElmAccessPoint.Count > 0)
            {
                if (_dbt == null)
                    _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                string msg = "'當天沒有執行交換步驟1，不執行步驟2，沒有回寫任何缺曠或成績資料。'";
                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use " + dbName);
                sb_SQL.AppendLine("insert into transfer_summary(begin_time,end_time,transfer_step,is_successful,summary) values('" + _beginDateTime.ToString("yyyy/MM/dd HH:mm:ss") + "','" + _endDateTime.ToString("yyyy/MM/dd HH:mm:ss") + "',2," + 0 + "," + msg + ")");


                _dbt.ExecuteNonQuerySQL(sb_SQL);

            }
        
        }


        public void WriteDBStep3()
        {
            _endDateTime = DateTime.Now;
            sb_SQL.Clear();
            if (_ElmAccessPoint.Count > 0)
            {
                if(_dbt == null )
                    _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                StringBuilder msg = new StringBuilder();
                msg.Append("'");
                foreach (string str in _dbt.GetData().GetAllDataDict().Keys)
                {
                    foreach (string val in _dbt.GetData().GetAllDataDict()[str])
                    {
                        msg.AppendLine(val);
                    }
                }
                msg.AppendLine("總共費時:"+ sp.Elapsed.Minutes + "分" + sp.Elapsed.Seconds + "秒");
                msg.Append("'");
                string dbName = _ElmAccessPoint[0].Attribute("DBName").Value;
                sb_SQL.AppendLine("use " + dbName);
                sb_SQL.AppendLine("insert into transfer_summary(begin_time,end_time,transfer_step,is_successful,summary) values('" + _beginDateTime.ToString("yyyy/MM/dd HH:mm:ss") + "','" + _endDateTime.ToString("yyyy/MM/dd HH:mm:ss") + "',3," + _IsProcessPass + "," + msg.ToString() + ")");
                

                _dbt.ExecuteNonQuerySQL(sb_SQL);
                    
            }        
        }


        /// <summary>
        /// 回寫成績到ischool
        /// </summary>
        public void InsertOrUpdateScoreToIschool()
        {
            try
            {
                int idx = 0;
                if (_dbt == null)
                    _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                foreach (XElement elmAccessPoint in _ConfMang.GetAccessPointList())
                {
                    QuizTransfer qt = new QuizTransfer(_ElmAccessPoint[idx].Attribute("Name").Value, _dbt, _ElmAccessPoint[idx].Attribute("AccessPoint").Value, _ElmAccessPoint[idx].Attribute("ContractName").Value, _ElmAccessPoint[idx].Attribute("UserName").Value, _ElmAccessPoint[idx].Attribute("Password").Value);
                    qt.WriteDataToischool();
                    idx++;
                }
            }
            catch (Exception ex)
            {
                Global._ExceptionManager.AddMessage(ex.Message);
            }
        
        }

        /// <summary>
        /// 將缺曠資料回寫到 ischool
        /// </summary>
        public void InsertOrUpdateAttendanceToIschool()
        {
            try
            {
                if (_dbt == null)
                    _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

                string ServiceName_GetSchoolYearSemester = "data.GetDefaultSchoolYearSemester";
                string ServiceName_InsertUpdateAttendance = "data.InsertUpdateAttendance";

                foreach (XElement elmAccessPoint in _ConfMang.GetAccessPointList())
                {
                    List<StudAttendanceRecord> AttendanceRec = new List<StudAttendanceRecord>();

                    XElement Req = new XElement("Request");
                    // 取得預設學年度學期
                    _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, ServiceName_GetSchoolYearSemester, Req);

                    XElement rsp = _dsat.GetResponse();
                    if (rsp != null)
                    {
                        Global._DefaultSchoolYear = rsp.Element("SystemConfig").Element("DefaultSchoolYear").Value;
                        Global._DefaultSemester = rsp.Element("SystemConfig").Element("DefaultSemester").Value;
                    }

                    sb_SQL.Clear();
                    sb_SQL.AppendLine("select student.id,class.dsa_source,attendance.occur_date,attendance.detail from attendance inner join student on student.login_name=attendance.ref_student_login_name inner join class on student.ref_class_name=class.class_name where class.dsa_source='" + elmAccessPoint.Attribute("Name").Value + "' and convert(varchar,last_update,1)=convert(varchar,getdate(),1) ");

                    DataTable dtAllInsert = _dbt.ExecuteQuerySQLDT(sb_SQL);

                    foreach (DataRow dr in dtAllInsert.Rows)
                    {
                        StudAttendanceRecord rec = new StudAttendanceRecord();
                        rec.StudentID = dr[0].ToString();
                        rec.SchoolYear = Global._DefaultSchoolYear;
                        rec.Semester = Global._DefaultSemester;
                        // 因為 Service 判斷新增或更新，日期需要補0格式2011/09/01
                        rec.OccurDate = DateTime.Parse(dr[2].ToString()).ToString("yyyy/MM/dd");
                        rec.Detail = dr[3].ToString();
                        AttendanceRec.Add(rec);
                    }

                    // 處理 XML DSA 

                    // 新增或更新
                    if (AttendanceRec.Count > 0)
                    {
                        //   <Attendance>                        
                        //        <!--以下為必要欄位-->
                        //        <RefStudentId>56006</RefStudentId>
                        //        <SchoolYear>100</SchoolYear>
                        //        <Semester>2</Semester>
                        //        <OccurDate>2011/8/30</OccurDate>
                        //        <Detail>a</Detail>                        
                        //</Attendance>
                        XElement elmInsert = new XElement("Request");
                        foreach (StudAttendanceRecord rec in AttendanceRec)
                        {
                            XElement attend = new XElement("Attendance");
                            attend.SetElementValue("RefStudentId", rec.StudentID);
                            attend.SetElementValue("SchoolYear", rec.SchoolYear);
                            attend.SetElementValue("Semester", rec.Semester);
                            attend.SetElementValue("OccurDate", rec.OccurDate);
                            attend.SetElementValue("Detail", rec.Detail);
                            elmInsert.Add(attend);
                        }
                        _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, ServiceName_InsertUpdateAttendance, elmInsert);
                        XElement rspData = _dsat.GetResponse();
                    }
                }
            }
            catch (Exception ex)
            {
                Global._ExceptionManager.AddMessage(ex.Message);
            }
        }


        ///// <summary>
        ///// 新增或更新學期成績
        ///// </summary>
        //public void InsertOrUpdateSC_AttendQuery()
        //{
        //    // 建立資料庫連線
        //    if (_dbt == null)
        //        _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

        //    // ServiceName
        //    string ServiceName_GetSchoolYearSemester = "data.GetDefaultSchoolYearSemester";
        //    string ServiceName_GetSCAttendList = "data.GetSCAttendList";

        //    // Store Procedure Name
        //    string StoreProcedureName = "sp_InsertOrUpdate_sc_attend";

        //    List<List<SqlParm>> sqlParameterList = new List<List<SqlParm>>();


        //    // 呼叫並執行 Store Procedure
        //    foreach (XElement elmAccessPoint in _ConfMang.GetAccessPointList())
        //    {
        //        List<StudAttendanceRecord> AttendanceRec = new List<StudAttendanceRecord>();

        //        XElement Req = new XElement("Request");
        //        // 取得預設學年度學期
        //        _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, ServiceName_GetSchoolYearSemester, Req);

        //        XElement rsp = _dsat.GetResponse();
        //        if (rsp != null)
        //        {
        //            Global._DefaultSchoolYear = rsp.Element("SystemConfig").Element("DefaultSchoolYear").Value;
        //            Global._DefaultSemester = rsp.Element("SystemConfig").Element("DefaultSemester").Value;
        //        }

        //        XElement Req_SC = new XElement("Request");
        //        Req_SC.SetElementValue("Condition", "");
        //        Req_SC.Element("Condition").SetElementValue("SchoolYear", Global._DefaultSchoolYear);
        //        Req_SC.Element("Condition").SetElementValue("Semester", Global._DefaultSemester);

        //        _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, ServiceName_GetSCAttendList, Req_SC);


        //        // 解析資料,StudentLoginName,CourseName,Score
        //        XElement rspElm = _dsat.GetResponse();

        //        foreach (XElement elm in rspElm.Elements("SCAttend"))
        //        {
        //            List<SqlParm> parms = new List<SqlParm>();
        //            SqlParm p1 = new SqlParm();
        //            p1.Type = "string";
        //            p1.Name = "@LoginName";
        //            p1.Value = elm.Element("StudentLoginName").Value;

        //            SqlParm p2 = new SqlParm();
        //            p2.Type = "string";
        //            p2.Name = "@CourseName";
        //            p2.Value = elm.Element("CourseName").Value;

        //            SqlParm p3 = new SqlParm();
        //            p3.Type = "number";
        //            p3.Name = "@Score";
        //            if (elm.Element("Score") == null || elm.Element("Score").Value == "")
        //                p3.Value = "null";
        //            else
        //                p3.Value = elm.Element("Score").Value;

        //            parms.Add(p1);
        //            parms.Add(p2);
        //            parms.Add(p3);

        //            sqlParameterList.Add(parms);

        //        }

        //        _dbt.ExecuteNonQueryStoreProcedure(sqlParameterList, StoreProcedureName, "sc_attend");

        //    }

        //}



        ///// <summary>
        ///// 新增或更新學期成績
        ///// </summary>
        //public void InsertOrUpdateSC_Attend()
        //{
        //    // 建立資料庫連線
        //    if (_dbt == null)
        //        _dbt = new DBTransfer(_ElmAccessPoint[0].Attribute("DBServerName").Value, _ElmAccessPoint[0].Attribute("DBName").Value, _ElmAccessPoint[0].Attribute("DBUserName").Value, _ElmAccessPoint[0].Attribute("DBPassword").Value);

        //    // ServiceName
        //    string ServiceName_GetSchoolYearSemester = "data.GetDefaultSchoolYearSemester";
        //    string ServiceName_GetSCAttendList = "data.GetSCAttendList";

        //    // Store Procedure Name
        //    string StoreProcedureName = "sp_InsertOrUpdate_sc_attend";

        //    List<List<SqlParm>> sqlParameterList = new List<List<SqlParm>>();
            

        //    // 呼叫並執行 Store Procedure
        //    foreach (XElement elmAccessPoint in _ConfMang.GetAccessPointList())
        //    {
        //        List<StudAttendanceRecord> AttendanceRec = new List<StudAttendanceRecord>();

        //        XElement Req = new XElement("Request");
        //        // 取得預設學年度學期
        //        _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, ServiceName_GetSchoolYearSemester, Req);

        //        XElement rsp = _dsat.GetResponse();
        //        if (rsp != null)
        //        {
        //            Global._DefaultSchoolYear = rsp.Element("SystemConfig").Element("DefaultSchoolYear").Value;
        //            Global._DefaultSemester = rsp.Element("SystemConfig").Element("DefaultSemester").Value;
        //        }

        //        XElement Req_SC = new XElement("Request");
        //        Req_SC.SetElementValue("Condition", "");
        //        Req_SC.Element("Condition").SetElementValue("SchoolYear", Global._DefaultSchoolYear);
        //        Req_SC.Element("Condition").SetElementValue("Semester", Global._DefaultSemester);

        //        _dsat = new DSATransfer(elmAccessPoint.Attribute("AccessPoint").Value, elmAccessPoint.Attribute("ContractName").Value, elmAccessPoint.Attribute("UserName").Value, elmAccessPoint.Attribute("Password").Value, ServiceName_GetSCAttendList, Req_SC);


        //        // 解析資料,StudentLoginName,CourseName,Score
        //        XElement rspElm = _dsat.GetResponse();
                
        //        foreach (XElement elm in rspElm.Elements("SCAttend"))
        //        {
        //            List<SqlParm> parms = new List<SqlParm>();
        //            SqlParm p1 = new SqlParm();
        //            p1.Type = "string";
        //            p1.Name  = "@LoginName";
        //            p1.Value = elm.Element("StudentLoginName").Value;

        //            SqlParm p2 = new SqlParm();
        //            p2.Type = "string";
        //            p2.Name = "@CourseName";
        //            p2.Value = elm.Element("CourseName").Value;

        //            SqlParm p3 = new SqlParm();
        //            p3.Type = "number";
        //            p3.Name = "@Score";
        //            if (elm.Element("Score") == null || elm.Element("Score").Value == "")
        //                p3.Value = "null";
        //            else
        //                p3.Value= elm.Element("Score").Value;

        //            parms.Add(p1);
        //            parms.Add(p2);
        //            parms.Add(p3);

        //            sqlParameterList.Add(parms);
                    
        //        }

        //        _dbt.ExecuteNonQueryStoreProcedure(sqlParameterList, StoreProcedureName,"sc_attend");

        //    }

        //}
    }
}
