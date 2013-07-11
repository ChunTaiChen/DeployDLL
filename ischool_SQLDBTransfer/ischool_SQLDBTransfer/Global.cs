using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ischool_SQLDBTransfer
{
    public class Global
    {
        /// <summary>
        /// DSA 目前服務存取點
        /// </summary>
        public static string _DSACurrectAccessPoint = "";
        /// <summary>
        /// DSA 目前存取名稱
        /// </summary>
        public static string _DSACurrectName = "";

        /// <summary>
        /// 預設學年度
        /// </summary>
        public static string _DefaultSchoolYear="";

        /// <summary>
        /// 預設學期
        /// </summary>
        public static string _DefaultSemester = "";

        public static ExceptionManager _ExceptionManager = new ExceptionManager();

    }
}
