using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 定期評量成績(資料交換XML用)
    /// </summary>
    public class SceTakeRecord
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
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 定期成績
        /// </summary>
        public string Score { get; set; }

        /// <summary>
        /// 平時成績
        /// </summary>
        public string AssigmentScore { get; set; }

        /// <summary>
        /// 文字評量
        /// </summary>
        public string Text { get; set; }
    }
}
