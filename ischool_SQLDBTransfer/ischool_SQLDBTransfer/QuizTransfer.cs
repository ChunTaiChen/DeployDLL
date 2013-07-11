using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using System.Net;
using System.IO;
using System.Windows.Forms;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 小考資料交換
    /// </summary>
    public class QuizTransfer
    {
        string _AccessPoint = "";
        string _ContractName = "";
        string _UserName = "";
        string _Password = "";
        string _DSADataSource = "";

        DBTransfer _DBTransfer;
        public QuizTransfer(string dsa_source,DBTransfer DBTrans, string AccessPoint, string ContractName, string UserName, string Password)
        {
            _DBTransfer = DBTrans;
            _AccessPoint = AccessPoint;
            _ContractName = ContractName;
            _UserName = UserName;
            _Password = Password;
            _DSADataSource = dsa_source;

        }

        /// <summary>
        /// 取得小考樣板資料
        /// </summary>
        /// <returns></returns>
        private XElement GetQuizItem()
        {
            /*
<Request>
   <CourseExtension CourseID="9178">
<Extension Name="GradeItem">
	<GradeItem>
		<Item ExamID="1" SubExamId="1" Name="小考一" Weight="1"/>
		<Item ExamID="1" SubExamId="2" Name="小考二" Weight="1"/>
		<Item ExamID="1" SubExamId="3" Name="小考三" Weight="1"/>
	</GradeItem>
</Extension>
</CourseExtension>
</Request>           
             
             */

            XElement retElm = new XElement("Request");            

            List<QuizItemRecord> QuizItemRecordList = _DBTransfer.ExecuteQuerySQLDTGetQuizItem(_DSADataSource);
            // 取得課程IDList
            List<string> CourseIDList = (from data in QuizItemRecordList select data.CourseID).Distinct().ToList();

            foreach (string str in CourseIDList)
            {
                XElement elmCourseEx = new XElement("CourseExtension");
                elmCourseEx.SetAttributeValue("CourseID", str);
                XElement elmExtension = new XElement("Extension");
                elmExtension.SetAttributeValue("Name", "GradeItem");
                XElement elmGradeItem = new XElement("GradeItem");
                foreach (QuizItemRecord quiz in QuizItemRecordList.Where(x => x.CourseID == str))
                {
                    XElement item = new XElement("Item");
                    item.SetAttributeValue("ExamID", quiz.ExamID);
                    item.SetAttributeValue("SubExamID", quiz.SubExamID);
                    item.SetAttributeValue("Name", quiz.Name);
                    item.SetAttributeValue("Weight", quiz.Weight);
                    elmGradeItem.Add(item);
                }
                elmExtension.Add(elmGradeItem);
                elmCourseEx.Add(elmExtension);
                retElm.Add(elmCourseEx);                
            }            
            return retElm;
        }

        /// <summary>
        /// 取得小考成績
        /// </summary>
        /// <returns></returns>
        private XElement GetQuizScore(List<QuizScoreRecord> QuizScoreRecordList,string CoID)
        {
            /*
Sample.SetExtensionDataMKNew
<Request> 
<SCAttendExtension CourseID="9178" StudentID="55038">
        <Extension Name="GradeBook">
          <Exam ExamID="6" Score="1">  <!-- 第一次定期評量下的小考  -->
            <Item SubExamID="1" Score="1"/>
            <Item SubExamID="2" Score=""/>
            <Item SubExamID="3" Score=""/>
            <Item SubExamID="4" Score=""/>
            <Item SubExamID="5" Score=""/>
            <Item SubExamID="6" Score=""/>
          </Exam>
          <Exam ExamID="7" Score="">   <!-- 第二次定期評量下的小考  -->
            <Item SubExamID="7" Score=""/>
            <Item SubExamID="8" Score=""/>
            <Item SubExamID="9" Score=""/>
            <Item SubExamID="10" Score=""/>
          </Exam>
          <Exam ExamID="8" Score="84"/> <!-- 第三次定期評量下的小考  -->
        </Extension>
      </SCAttendExtension>
</Request> 
             */

            XElement retElm = new XElement("Request");            

            //List<QuizScoreRecord> QuizScoreRecordList = _DBTransfer.ExecuteQuerySQLDTGetQuizScore(_DSADataSource);

            //// 取得課程ID
            //List<string> CousreIDList = (from data in QuizScoreRecordList select data.CourseID).Distinct().ToList();

            // 取得學生ID
            List<string> StudentIDList = (from data in QuizScoreRecordList select data.StudentID).Distinct().ToList();

            // 取得試別ID
            List<string> ExamIDList = (from data in QuizScoreRecordList select data.ExamID).Distinct().ToList();

            List<QuizScoreRecord> QuizScoreRecordTmp = new List<QuizScoreRecord>();

            int dataCount = 0,subExamCount=0;

            foreach (string StudID in StudentIDList)
            {
                //foreach (string CoID in CousreIDList)
                //{
                    dataCount = 0;
                        QuizScoreRecordTmp=(from data in QuizScoreRecordList where data.StudentID == StudID && data.CourseID == CoID select data).ToList();
                        XElement elmExtension = new XElement("Extension");
                        elmExtension.SetAttributeValue("Name", "GradeBook");
                        foreach (string ExamID in ExamIDList)
                        {
                            subExamCount = 0;
                            XElement elmExam = new XElement("Exam");
                            foreach (QuizScoreRecord qsr in QuizScoreRecordTmp.Where(x=>x.ExamID==ExamID).ToList())
                            {                                
                                XElement elmItem = new XElement("Item");
                                elmItem.SetAttributeValue("SubExamID", qsr.SubExamID);
                                elmItem.SetAttributeValue("Score", qsr.SubScore);
                                elmExam.Add(elmItem);                                
                                dataCount++;
                                subExamCount++;
                            }
                            if (subExamCount > 0)
                            {
                                elmExam.SetAttributeValue("ExamID", ExamID);
                                elmExtension.Add(elmExam);
                            }
                        }
                        if (dataCount > 0)
                        {
                            XElement elmSCAttendExtension = new XElement("SCAttendExtension");
                            elmSCAttendExtension.SetAttributeValue("CourseID", CoID);
                            elmSCAttendExtension.SetAttributeValue("StudentID",StudID);

                            elmSCAttendExtension.Add(elmExtension);
                            retElm.Add(elmSCAttendExtension);
                        }
                 //}
            }
            return retElm;
        }

        /// <summary>
        /// 取得定期評量成績
        /// </summary>
        /// <returns></returns>
        private XElement GetSceTake()
        {
            /*
             Sample.SetCourseExamScoreWithExtensionNew
<Course CourseID="11737">
	<Exam ExamID="61">
		<Student StudentID="58444" Score="0">
			<Extension>
				<Score>99</Score>
				<AssignmentScore>100</AssignmentScore>
				<Text>表現優良</Text>
			</Extension>
		</Student>
	</Exam>
</Course>             
             */

            XElement retElm = new XElement("Request");
            List<SceTakeRecord> SceTakeRecordList = _DBTransfer.ExecuteQuerySQLDTGetSceTake(_DSADataSource);

            // 取得課程編號
            List<string> CourseIDList = (from data in SceTakeRecordList select data.CourseID).Distinct().ToList();
            
            // 取得試別編號
            List<string> ExamIDList = (from data in SceTakeRecordList select data.ExamID).Distinct().ToList();
            int CoCount = 0,ExCount = 0;
            foreach (string CourseID in CourseIDList)
            {
                CoCount = 0;
                XElement elmCourse = new XElement("Course");
                foreach (string ExamID in ExamIDList)
                {
                    ExCount = 0;
                    XElement elmExam = new XElement("Exam");
                    
                    foreach (SceTakeRecord str in (from data in SceTakeRecordList where data.ExamID==ExamID && data.CourseID== CourseID select data).ToList())
                    {                           
                        XElement elmStudent = new XElement("Student");
                        elmStudent.SetAttributeValue("StudentID", str.StudentID);
                        if(string.IsNullOrEmpty(str.Score))
                            elmStudent.SetAttributeValue("Score","0");
                        else
                            elmStudent.SetAttributeValue("Score", str.Score);
                        
                        XElement elmExtension = new XElement("Extension");
                        XElement elmExtensionS = new XElement("Extension");
                        elmExtensionS.SetElementValue("Score", str.Score);
                        elmExtensionS.SetElementValue("AssignmentScore", str.AssigmentScore);
                        elmExtensionS.SetElementValue("Text", str.Text);
                        elmExtension.Add(elmExtensionS);
                        elmStudent.Add(elmExtension);
                        elmExam.Add(elmStudent);
                        CoCount++;
                        ExCount++;
                    }

                    if (ExCount > 0)
                    {
                        elmExam.SetAttributeValue("ExamID", ExamID);
                        elmCourse.Add(elmExam);
                    }

                }

                if (CoCount > 0)
                {                    
                    elmCourse.SetAttributeValue("CourseID", CourseID);
                    retElm.Add(elmCourse);
                }
            }
            return retElm;        
        }

        /// <summary>
        /// 回寫資料到 ischool
        /// </summary>
        /// <returns></returns>
        public string WriteDataToischool()
        {
            string retVal = "";

            List<XElement> elmCourseScoreList = new List<XElement> ();            
            DSATransfer Transfer;
            XElement el=null;
            try
            {

                // 小考樣板
                Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetExtensionDataNew", GetQuizItem());
                Transfer.GetResponse();

                // 小考成績
                List<QuizScoreRecord> QuizScoreRecordList = _DBTransfer.ExecuteQuerySQLDTGetQuizScore(_DSADataSource);
                // 取得課程ID
                List<string> CousreIDList = (from data in QuizScoreRecordList select data.CourseID).Distinct().ToList();
                // 小考分批
                foreach (string cid in CousreIDList)
                {
                    try
                    {
                        Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetExtensionDataMKNew", GetQuizScore(QuizScoreRecordList, cid));
                        Transfer.GetResponse();
                    }
                    catch (WebException webxs1)
                    {
                        Global._ExceptionManager.AddMessage("try 2:" + FISCA.ErrorReport.Generate(webxs1));
                        Thread.Sleep(10000);
                        try
                        {
                            Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetExtensionDataMKNew", GetQuizScore(QuizScoreRecordList, cid));
                            Transfer.GetResponse();
                        }
                        catch (WebException webxs2)
                        {
                            Global._ExceptionManager.AddMessage("try 1:" + FISCA.ErrorReport.Generate(webxs2));
                            Thread.Sleep(10000);
                            Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetExtensionDataMKNew", GetQuizScore(QuizScoreRecordList, cid));
                            Transfer.GetResponse();
                        }
                    
                    }

                }

                // 定期評量
                foreach (XElement elm in GetSceTake().Elements("Course"))
                    elmCourseScoreList.Add(elm);

                foreach (XElement elm in elmCourseScoreList)
                {
                    try
                    {
                        el = elm;
                        Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetCourseExamScoreWithExtensionNew", elm);
                        Transfer.GetResponse();
                    }
                    catch (WebException webex)
                    {
                        Global._ExceptionManager.AddMessage("try 2:" + FISCA.ErrorReport.Generate(webex));
                        Thread.Sleep(10000);
                        try
                        {
                            Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetCourseExamScoreWithExtensionNew", elm);
                            Transfer.GetResponse();
                        }
                        catch (WebException webEx1)
                        {
                            Global._ExceptionManager.AddMessage("try 1:" + FISCA.ErrorReport.Generate(webEx1));
                            Thread.Sleep(10000);
                            Transfer = new DSATransfer(_AccessPoint, _ContractName, _UserName, _Password, "data.SetCourseExamScoreWithExtensionNew", elm);
                            Transfer.GetResponse();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string fileName = "ResponseError" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".txt";
                string file_path = Application.StartupPath + "\\" + fileName;
                StreamWriter sw = File.CreateText(file_path);
                sw.Write(FISCA.ErrorReport.Generate(ex));
                sw.Close();

                Global._ExceptionManager.AddMessage(_AccessPoint + "," + _ContractName + "," + ex.Message + "," + el.ToString()+","+FISCA.ErrorReport.Generate(ex));             
                             
            }
            
            return retVal;
        }
    }
}
