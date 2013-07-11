using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 小考成績(資料交換XML用)
    /// </summary>
    public class QuizScoreRecord
    {
        /// <summary>
        /// 課程編號
        /// </summary>
        public string CourseID { get; set; }

        /// <summary>
        /// 學生編號
        /// </summary>
        public string StudentID { get; set; }
        
        /// <summary>
        /// 試別編號
        /// </summary>
        public string ExamID { get; set; }

        /// <summary>
        /// 小考平均成績
        /// </summary>
        public string Score { get; set; }

        /// <summary>
        /// 小考編號
        /// </summary>
        public string SubExamID { get; set; }

        /// <summary>
        /// 小考成績
        /// </summary>
        public string SubScore { get; set; }
    }
}
