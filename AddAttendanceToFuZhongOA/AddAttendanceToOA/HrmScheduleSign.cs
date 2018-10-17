using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddAttendanceToOA
{
    /// <summary>
    /// OA考勤对象
    /// </summary>
    class HrmScheduleSign
    {
        /// <summary>
        /// 用户Id
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// 用户姓名
        /// </summary>
        public string UserName { get; set; }
        /// <summary>
        /// 打卡类型(1:上班,2:下班)
        /// </summary>
        public string SignType {get;set;}
        /// <summary>
        /// 签到日期
        /// </summary>
        public string SignDate { get; set; }
        /// <summary>
        /// 签到时间
        /// </summary>
        public string SignTime { get; set; }

        /// <summary>
        /// 考勤机编号
        /// </summary>
        public string WorkCode { get; set; }
    }

    class AttendDuty
    {
        /// <summary>
        /// 用户Id
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// 用户姓名
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// 签到日期
        /// </summary>
        public string SignDate { get; set; }

        /// <summary>
        /// 签到时间
        /// </summary>
        public long SignInTime { get; set; }

        /// <summary>
        /// 签退时间
        /// </summary>
        public long SignOutTime { get; set; }

        /// <summary>
        /// 工号
        /// </summary>
        public string WorkCode { get; set; }
    }
}
