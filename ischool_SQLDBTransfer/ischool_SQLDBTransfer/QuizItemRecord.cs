using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 小考樣板紀錄(資料交換XML用)
    /// </summary>
    public class QuizItemRecord
    {
        /// <summary>
        /// 課程編號
        /// </summary>
        public string CourseID { get; set; }


        /// <summary>
        /// 試別編號
        /// </summary>
        public string ExamID { get; set; }


        /// <summary>
        /// 小考編號
        /// </summary>
        public string SubExamID { get; set; }

        /// <summary>
        /// 小考名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 比重
        /// </summary>
        public string Weight { get; set; }
    }
}
