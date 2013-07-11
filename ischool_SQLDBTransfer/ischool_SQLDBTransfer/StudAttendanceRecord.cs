using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ischool_SQLDBTransfer
{
    /// <summary>
    /// 學生缺況紀錄資料交換用
    /// </summary>
    class StudAttendanceRecord
    {
        /// <summary>
        /// 學生編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public string Semester { get; set; }

        /// <summary>
        /// 缺曠日期
        /// </summary>
        public string OccurDate { get; set; }

        /// <summary>
        /// 缺曠明細
        /// </summary>
        public string Detail { get; set; }
    }
}
