using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 與 SQL Server 交換
    /// </summary>
    public class DBTransfer
    {
        string _connectStr="";
        SqlTransaction TranExecuteNonQuery;
        SqlTransaction TranExecuteQuery;
        SqlConnection cn;
        SqlCommand cmd;
        DataStore _DataStore;
        

        /// <summary>
        /// 建立連接字串
        /// </summary>
        /// <param name="DBServer"></param>
        /// <param name="DBName"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        public DBTransfer(string DBServer,string DBName,string UserName,string Password)
        {
            _connectStr = "Data Source=" + DBServer + ";Initial Catalog=" + DBName + ";uid=" + UserName + ";pwd=" + Password + ";";
            cn = new SqlConnection();
            cmd = new SqlCommand();
            _DataStore = new DataStore();
        }

        /// <summary>
        /// 執行SQL Query(沒有回傳資料)
        /// </summary>
        /// <param name="Query"></param>
        /// <returns></returns>
        public bool ExecuteNonQuerySQL(StringBuilder Query)
        {
            bool pass = true;
            cn.ConnectionString = _connectStr;
            try
            {
                cn.Open();
                TranExecuteNonQuery = cn.BeginTransaction("Query start");
                cmd.Transaction = TranExecuteNonQuery;
                cmd.Connection = cn;
                cmd.CommandText = Query.ToString();
                cmd.ExecuteNonQuery();
                TranExecuteNonQuery.Commit();
            }
            catch
            {
                TranExecuteNonQuery.Rollback();
                pass = false;
            }
            finally
            {
                cn.Close();
            }
            return pass;
        }

        /// <summary>
        /// 執行SQL Query(有回傳資料)
        /// </summary>
        /// <param name="Query"></param>
        /// <returns></returns>
        public List<string> ExecuteQuerySQL(StringBuilder Query)
        {
            List<string> retVal = new List<string>();
            cn.ConnectionString = _connectStr;
            try
            {
                cn.Open();
                TranExecuteQuery = cn.BeginTransaction("Query start");
                cmd.Transaction = TranExecuteQuery;
                cmd.Connection = cn;


                cmd.CommandText = Query.ToString();
                SqlDataReader Reader = cmd.ExecuteReader();
                while (Reader.Read())
                {
                    retVal.Add(Reader[0].ToString());
                }
                Reader.Close();

                TranExecuteQuery.Commit();
                cn.Close();
            }
            catch
            {
                TranExecuteQuery.Rollback();
            }
            finally
            {
                cn.Close();
            }

            return retVal;
        }


        /// <summary>
        /// 執行SQL Query(有回傳資料)
        /// </summary>
        /// <param name="Query"></param>
        /// <returns></returns>
        public DataTable ExecuteQuerySQLDT(StringBuilder Query)        
        {
            DataTable dt = new DataTable();
            cn.ConnectionString = _connectStr;
            try
            {
                cn.Open();
                TranExecuteQuery = cn.BeginTransaction("Query start");
                cmd.Transaction = TranExecuteQuery;
                cmd.Connection = cn;

                cmd.CommandText = Query.ToString();
                SqlDataReader Reader = cmd.ExecuteReader();
                dt.Load(Reader);
                Reader.Close();
                TranExecuteQuery.Commit();
                cn.Close();
            }
            catch
            {
                TranExecuteQuery.Rollback();
            }
            finally
            {
                cn.Close();
            }

            return dt;
        }
        
        /// <summary>
        /// 執行SQL Query(傳入 DataStore 方式)
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        public bool ExecuteQuerySQLByDataStore(DataStore ds)
        {
            bool pass = true;
            int DataBeginCount = 0;
            int DataTableCount = 0;
            int ExeCount = 0;
            cn.ConnectionString = _connectStr;
            cn.Open();    
            TranExecuteQuery = cn.BeginTransaction("Query start");
            cmd.Transaction = TranExecuteQuery;
            cmd.Connection = cn;
            string name = "";

            try
            {
                foreach (string str in ds.GetAllDataDict().Keys)
                {
                    name = str;
                    DataBeginCount = ds.GetDataListCountByName(str);
                    ExeCount = 0;

                    foreach (string ComStr in ds.GetAllDataDict()[str])
                    {
                        cmd.CommandText = ComStr;
                        cmd.ExecuteNonQuery();
                        ExeCount++;
                    }

                    if(str=="student_update")
                        cmd.CommandText = "select count(*) from student";
                    else
                        cmd.CommandText = "select count(*) from "+str;
                    DataTableCount=0;
                    int tempInt;
                    if(int.TryParse(cmd.ExecuteScalar().ToString (),out tempInt ))
                        DataTableCount=tempInt;

                    _DataStore.AddData(str, str + ":Service取得資料共有:" + DataBeginCount + "筆,轉換寫入資料庫:" + ExeCount + "筆,資料庫內共有"+DataTableCount+"筆.");
                }

                TranExecuteQuery.Commit();
            }

            catch (Exception ex)
            {
                _DataStore.AddData(name, name+":"+ex.Message);
                pass = false;
                Global._ExceptionManager.AddMessage(name + ":" + ex.Message);
                TranExecuteQuery.Rollback();
            }
                finally
                {
                    cn.Close();
                }
            

            return pass ;
        }

        /// <summary>
        /// 取得儲存的訊息資料
        /// </summary>
        /// <returns></returns>
        public DataStore GetData()
        {
            return _DataStore;
        }

        /// <summary>
        /// 執行傳入參數的Store Procedure
        /// </summary>
        /// <param name="sqlParameterDict"></param>
        /// <returns></returns>
        public bool ExecuteNonQueryStoreProcedure(List<List<SqlParm>> sqlParameterList,string StoreProcedureName,string tableName)
        {
            bool pass = true;
            cn.ConnectionString = _connectStr;
            int count = 0;
            try
            {
                cn.Open();
                TranExecuteNonQuery = cn.BeginTransaction("Query start");
                cmd.CommandType = CommandType.StoredProcedure;                
                cmd.Transaction = TranExecuteNonQuery;
                cmd.Connection = cn;
                cmd.CommandText = StoreProcedureName;
                foreach (List<SqlParm> parm in sqlParameterList)
                {
                    cmd.Parameters.Clear();
                    foreach(SqlParm p in parm)
                    {
                        SqlParameter p1 = new SqlParameter();
                        p1.ParameterName = p.Name;
                        if (p.Type == "string")
                            p1.SqlDbType = SqlDbType.NVarChar;
                        else
                            p1.SqlDbType = SqlDbType.Real;

                        if (p.Value == "null")
                            p1.SqlValue = DBNull.Value;
                        else
                            p1.SqlValue = p.Value;
                        
                        cmd.Parameters.Add(p1);
                        
                    
                    }
                    count++;
                    cmd.ExecuteNonQuery();
                }
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select count(*) from " + tableName;
                int DataTableCount = 0;
                int tempInt;
                if (int.TryParse(cmd.ExecuteScalar().ToString(), out tempInt))
                    DataTableCount = tempInt;

                _DataStore.AddData(StoreProcedureName, StoreProcedureName + ":程式取得資料共有:" + sqlParameterList.Count + "筆,資料庫內資料表:"+tableName+",共有" + DataTableCount + "筆.");
                TranExecuteNonQuery.Commit();
            }
            catch(Exception ex)
            {
                Global._ExceptionManager.AddMessage(ex.Message);
                TranExecuteNonQuery.Rollback();
                pass = false;
            }
            finally
            {
                cn.Close();
            }
            return pass;
        }

        /// <summary>
        /// 特別處理sc_attend差異用
        /// </summary>
        /// <returns></returns>
        public bool ExecuteSpc_SC_Attend()
        {
            bool pass = true;            
            int DataTableCount = 0;
            int updateCount = 0, insertCount = 0;
            cn.ConnectionString = _connectStr;
            if (cmd == null)
                cmd = new SqlCommand();

            cn.Open();
            TranExecuteQuery = cn.BeginTransaction("Query start");
            cmd.Transaction = TranExecuteQuery;
            cmd.Connection = cn;
            string tableName = "sc_attend";           

            try
            {
                // 更新 Query               
                // 取得須更新筆數
                cmd.CommandText = "select count(*) from sc_attend_ischool inner join sc_attend on sc_attend_ischool.ref_student_login_name=sc_attend.ref_student_login_name and sc_attend_ischool.ref_course_course_name=sc_attend.ref_course_course_name;";
                int tempInt;
                string str=cmd.ExecuteScalar().ToString();
                if (int.TryParse(str, out tempInt))
                    updateCount = tempInt;

                // 執行更新
                cmd.CommandText = "update sc_attend set sc_attend.ischool_score=sc_attend_ischool.ischool_score  from sc_attend_ischool inner join sc_attend on sc_attend_ischool.ref_student_login_name=sc_attend.ref_student_login_name and sc_attend_ischool.ref_course_course_name=sc_attend.ref_course_course_name;";
                cmd.ExecuteNonQuery();
                
                // 新增 Query
                // 取得需要新增筆數
                cmd.CommandText = "select count(*) from sc_attend_ischool left join sc_attend on sc_attend_ischool.ref_student_login_name=sc_attend.ref_student_login_name and sc_attend_ischool.ref_course_course_name=sc_attend.ref_course_course_name where sc_attend.ref_student_login_name is null and sc_attend.ref_course_course_name is null;";
                str = cmd.ExecuteScalar().ToString();
                if (int.TryParse(str, out tempInt))
                    insertCount = tempInt;

                // 執行新增
                cmd.CommandText = "insert into sc_attend select sc_attend_ischool.* from sc_attend_ischool left join sc_attend on sc_attend_ischool.ref_student_login_name=sc_attend.ref_student_login_name and sc_attend_ischool.ref_course_course_name=sc_attend.ref_course_course_name where sc_attend.ref_student_login_name is null and sc_attend.ref_course_course_name is null;";
                cmd.ExecuteNonQuery();

                // 統計sc_attend總共筆數
                cmd.CommandText = "select count(*) from " + tableName;
                    DataTableCount = 0;
                    str = cmd.ExecuteScalar().ToString();     
                    if (int.TryParse(str, out tempInt))
                        DataTableCount = tempInt;

                    _DataStore.AddData(tableName, tableName + ":資料庫內共有" + DataTableCount + "筆,新增"+insertCount+"筆,更新"+updateCount+"筆.");

                TranExecuteQuery.Commit();
            }

            catch (Exception ex)
            {
                _DataStore.AddData(tableName, tableName+":"+ex.Message);
                Global._ExceptionManager.AddMessage(tableName + ":" + ex.Message);
                pass = false;
                TranExecuteQuery.Rollback();
            }
            finally
            {
                cn.Close();
            }

            return pass;
        }

        /// <summary>
        /// 取得小考樣板項目紀錄
        /// </summary>
        /// <returns></returns>
        public List<QuizItemRecord> ExecuteQuerySQLDTGetQuizItem(string DSADataSource)        
        {
            List<QuizItemRecord> retVal = new List<QuizItemRecord>();
            string TitleName = "小考樣板資料" +"_" +DSADataSource;
            string Query = "select course.id as CourseID,course_exam.ref_exam_id as ExamID,quiz_item.id as SubExamID,quiz_item.quiz_name as QuizName,quiz_item.weight as Weight from quiz_item inner join course_exam on quiz_item.ref_course_name=course_exam.ref_course_name and quiz_item.ref_exam_name=course_exam.exam_name inner join course on quiz_item.ref_course_name=course.course_name where course.dsa_source='" + DSADataSource + "';";
            cn.ConnectionString = _connectStr;
            try
            {
                // CourseID,ExamID,SubExamID,QuizName,Weight
                cn.Open();
                TranExecuteQuery = cn.BeginTransaction("Query start");
                cmd.Transaction = TranExecuteQuery;
                cmd.Connection = cn;

                cmd.CommandText = Query;
                SqlDataReader Reader = cmd.ExecuteReader();
                while (Reader.Read())
                {
                    QuizItemRecord qir = new QuizItemRecord();
                    qir.CourseID = Reader[0].ToString();
                    qir.ExamID = Reader[1].ToString();
                    qir.SubExamID = Reader[2].ToString();
                    qir.Name = Reader[3].ToString();
                    qir.Weight = Reader[4].ToString();
                    retVal.Add(qir);
                
                }
                Reader.Close();
                TranExecuteQuery.Commit();
                _DataStore.AddData(TitleName, TitleName + ":程式取得資料庫內共有:" + retVal.Count + "筆.");                
                cn.Close();
            }
            catch(Exception ex)
            {
                _DataStore.AddData(TitleName, TitleName + ":" + ex.Message);
                Global._ExceptionManager.AddMessage(TitleName + ":" + ex.Message);
                TranExecuteQuery.Rollback();
            }
            finally
            {
                cn.Close();
            }
            return retVal;
        }

        /// <summary>
        /// 取得小考成績
        /// </summary>
        /// <returns></returns>
        public List<QuizScoreRecord> ExecuteQuerySQLDTGetQuizScore(string DSADataSource)        
        {
            List<QuizScoreRecord> retVal = new List<QuizScoreRecord>();
            string TitleName = "小考成績資料" + "_" + DSADataSource;
            string Query = "select course.id as CourseID,student.id as StudentID,course_exam.ref_exam_id as ExamID,quiz_item.id as SubExamID,quiz_score.score as SubScore from quiz_item inner join course_exam on quiz_item.ref_course_name=course_exam.ref_course_name and quiz_item.ref_exam_name=course_exam.exam_name inner join course on quiz_item.ref_course_name=course.course_name inner join quiz_score on quiz_item.id=quiz_score.ref_quiz_item_id inner join student on quiz_score.ref_student_login_name=student.login_name inner join sc_attend_ischool on sc_attend_ischool.ref_student_login_name=student.login_name and sc_attend_ischool.ref_course_course_name=course.course_name where course.dsa_source='" + DSADataSource + "';";
            cn.ConnectionString = _connectStr;
            try
            {
                cn.Open();
                TranExecuteQuery = cn.BeginTransaction("Query start");
                cmd.Transaction = TranExecuteQuery;
                cmd.Connection = cn;
                // CourseID,StudentID,ExamID,SubExamID,SubScore
                cmd.CommandText = Query.ToString();
                SqlDataReader Reader = cmd.ExecuteReader();
                while (Reader.Read())
                {
                    QuizScoreRecord qsr = new QuizScoreRecord();
                    qsr.CourseID = Reader[0].ToString();
                    qsr.StudentID = Reader[1].ToString();
                    qsr.ExamID = Reader[2].ToString();
                    qsr.SubExamID = Reader[3].ToString();
                    qsr.SubScore = Reader[4].ToString();
                    retVal.Add(qsr);
                }
                Reader.Close();
                TranExecuteQuery.Commit();
                _DataStore.AddData(TitleName, TitleName + ":程式取得資料庫內共有:" + retVal.Count + "筆.");
                cn.Close();
            }
            catch(Exception ex)
            {
                _DataStore.AddData(TitleName, TitleName + ":" + ex.Message);
                Global._ExceptionManager.AddMessage(TitleName + ":" + ex.Message);
                TranExecuteQuery.Rollback();
            }
            finally
            {
                cn.Close();
            }
            return retVal;
        }


        /// <summary>
        /// 取得定期評量
        /// </summary>
        /// <returns></returns>
        public List<SceTakeRecord> ExecuteQuerySQLDTGetSceTake(string DSADataSource)        
        {
            List<SceTakeRecord> retVal = new List<SceTakeRecord>();
            string TitleName = "定期評量成績資料" + "_" + DSADataSource;
            string Query = "select course.id as CourseID,course_exam.ref_exam_id as ExamID,student.id as StudentID,sce_take.score as Score,sce_take.quiz_score as QuizScore,sce_take.text_score as TextScore from sce_take inner join student on sce_take.ref_student_login_name=student.login_name inner join course_exam on sce_take.ref_course_course_name=course_exam.ref_course_name and sce_take.ref_exam_name=course_exam.exam_name inner join course on course.course_name=sce_take.ref_course_course_name inner join sc_attend_ischool on sc_attend_ischool.ref_student_login_name=student.login_name and sc_attend_ischool.ref_course_course_name=course.course_name where course.dsa_source='" + DSADataSource + "';";
            cn.ConnectionString = _connectStr;
            try
            {
                cn.Open();
                TranExecuteQuery = cn.BeginTransaction("Query start");
                cmd.Transaction = TranExecuteQuery;
                cmd.Connection = cn;
                // CourseID,ExamID,StudentID,Score,QuizScore,TextScore
                cmd.CommandText = Query.ToString();
                SqlDataReader Reader = cmd.ExecuteReader();
                while (Reader.Read())
                {
                    SceTakeRecord str = new SceTakeRecord();
                    str.CourseID = Reader[0].ToString();
                    str.ExamID = Reader[1].ToString();
                    str.StudentID = Reader[2].ToString();
                    str.Score = Reader[3].ToString();
                    str.AssigmentScore = Reader[4].ToString();
                    str.Text = Reader[5].ToString();
                    retVal.Add(str);
                }
                Reader.Close();
                TranExecuteQuery.Commit();
                _DataStore.AddData(TitleName, TitleName + ":程式取得資料庫內共有:" + retVal.Count + "筆.");
                cn.Close();
            }
            catch(Exception ex)
            {
                _DataStore.AddData(TitleName, TitleName + ":" + ex.Message);
                Global._ExceptionManager.AddMessage(TitleName + ":" + ex.Message);
                TranExecuteQuery.Rollback();
            }
            finally
            {
                cn.Close();
            }

            return retVal;
        }
    }
}
