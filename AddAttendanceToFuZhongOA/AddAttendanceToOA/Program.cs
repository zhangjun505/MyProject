using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Linq;
namespace AddAttendanceToOA
{
    class Program
    {
        static MySqlConnection conn; //数据库连接对象
        static string attendanceDir = ""; //外地日考勤记录存放目录
        static string[] arrExcelType; //考勤文件类别数组
        static void Main(string[] args)
        {
            if (Init())
            {
                string logFileName = Application.StartupPath + "\\log.txt";
                FileStream fs;
                if (!File.Exists(logFileName))
                    fs = new FileStream(logFileName, FileMode.Create);
                else
                    fs = new FileStream(logFileName, FileMode.Append);
                StreamWriter sw = new StreamWriter(fs);
                Console.WriteLine("正在同步考勤记录，请稍等......");
                string retMsg = ImportAttendanceFromExcel();
                string logInfo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "同步考勤记录:\r\n" + retMsg + "";
                sw.WriteLine(logInfo);
                sw.Close();
                fs.Close();
                Console.WriteLine(logInfo);
            }
        }

        /// <summary>
        /// 初始化
        /// </summary>
        static bool Init()
        {
            if (ConfigurationManager.ConnectionStrings["OADbConn"] == null)
            {
                Console.WriteLine("配置文件【connectionStrings】配置节中缺少【OADbConn】配置项");
                return false;
            }
            else
            {
                conn = new MySqlConnection(ConfigurationManager.ConnectionStrings["OADbConn"].ConnectionString);
                try
                {
                    DBHelper.OpenDB(conn);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("数据库打开失败");
                    return false;
                }
            }
            if (ConfigurationManager.AppSettings["AttendanceDir"] == null)
            {
                Console.WriteLine("配置文件【appSettings】配置节中缺少【AttendanceDir】配置项");
                return false;
            }
            else
            {
                attendanceDir = ConfigurationManager.AppSettings["AttendanceDir"];
                if (!Directory.Exists(attendanceDir))
                {
                    Console.WriteLine("外地日考勤目录【" + attendanceDir+"】不存在！");
                    return false;
                }
            }
            if (ConfigurationManager.AppSettings["ExcelType"] == null)
            {
                Console.WriteLine("配置文件【appSettings】配置节中缺少【ExcelType】配置项");
                return false;
            }
            else
            {
                arrExcelType = ConfigurationManager.AppSettings["ExcelType"].Split(';');
            }
            return true;
        }

        /// <summary>
        /// 将考勤记录集合转换为MySQL中的考勤表记录集合
        /// </summary>
        /// <param name="lstHrmScheduleSign">考勤记录集合</param>
        /// <returns></returns>
        static List<AttendDuty> ConvertScheduleSignToAttendDuty(List<HrmScheduleSign> lstHrmScheduleSign)
        {
            List<AttendDuty> lstAttendDuty=new List<AttendDuty>();
            var lstGroupScheduleSign =from schedule in lstHrmScheduleSign
                                      group schedule by new {schedule.UserName,schedule.WorkCode,schedule.SignDate} into g
                                      select new { UserName = g.Key.UserName, WorkCode = g.Key.WorkCode,SignDate = g.Key.SignDate, SignInTime = g.Min(p => p.SignTime), SignOutTime = g.Max(p => p.SignTime) };
            foreach (var groupScheduleSign in lstGroupScheduleSign)
            {
                AttendDuty attendDuty = new AttendDuty();
                attendDuty.UserName = groupScheduleSign.UserName;
                attendDuty.SignDate = groupScheduleSign.SignDate;
                attendDuty.SignInTime = !String.IsNullOrEmpty(groupScheduleSign.SignInTime)?ConvertDateTimeToInt(DateTime.Parse(groupScheduleSign.SignInTime)):0;
                attendDuty.SignOutTime = !String.IsNullOrEmpty(groupScheduleSign.SignOutTime)? ConvertDateTimeToInt(DateTime.Parse(groupScheduleSign.SignOutTime)) : 0;
                attendDuty.WorkCode = groupScheduleSign.WorkCode;
                lstAttendDuty.Add(attendDuty);
            }
            return lstAttendDuty;
        }
        /// <summary>
        /// 添加考勤记录到OA
        /// </summary>
        /// <param name="lstHrmScheduleSign">考勤记录集合</param>
        static int AddAttendance(List<HrmScheduleSign> lstHrmScheduleSign)
        {
            int recordCount = 0;
            string sql = "";
            lstHrmScheduleSign = lstHrmScheduleSign.OrderBy(p => p.UserName).ThenBy(p => p.SignDate).ToList();
            List<AttendDuty> lstAttendDuty = ConvertScheduleSignToAttendDuty(lstHrmScheduleSign);
            foreach (AttendDuty attendDuty in lstAttendDuty)
            {
                string workCode = "";
                string userId = GetUserId(attendDuty.UserName, out workCode);
                if (userId!="")
                {
                    attendDuty.UserId = userId;
                    if (string.IsNullOrEmpty(workCode) && !string.IsNullOrEmpty(attendDuty.WorkCode))
                    {
                        sql = "update user set job_number='" + attendDuty.WorkCode.PadLeft(6).Replace(" ","0") + "' where user_id='" + userId+"'";
                        DBHelper.ExecuteCommand(sql);
                    }
                    string signDate = attendDuty.SignDate;
                    sql = "delete from attend_duty where user_id='" + userId + "' and sign_date='" + signDate + "'";
                    DBHelper.ExecuteCommand(sql);

                    string newId = DBHelper.GetNewId("attend_duty", "Duty_Id");
                    sql = "insert into attend_duty(Duty_Id,User_Id,Sign_In_Time,Sign_In_Normal,Sign_Out_Time,Sign_Out_Normal,Duty_Type,Sign_Date) values(" + newId + ",'" + attendDuty.UserId
                        + "'," + attendDuty.SignInTime + ",1486602000," + attendDuty.SignOutTime + ",1487066400,2,'" + attendDuty.SignDate + "')";
                    int ret = DBHelper.ExecuteCommand(sql);
                    if (ret>0)
                    recordCount++;

                }
            }
            return recordCount;
        }

        /// <summary>  
        /// 将DateTime时间格式转换为Unix时间戳格式  
        /// </summary>  
        /// <param name="time">时间</param>  
        /// <returns>long</returns>  
        public static long ConvertDateTimeToInt(DateTime time)
        {
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1, 0, 0, 0, 0));
            long t = (time.Ticks - startTime.Ticks) / 10000000;   //除10000000调整为10位      
            return t;
        }

        /// <summary>
        /// 获取用户Id
        /// </summary>
        /// <param name="userName">用户姓名</param>
        /// <returns></returns>
        static string GetUserId(string userName,out string workCode)
        {
            string sql = "select concat(concat(convert(user_id,char),'/'),job_number) from user where user_name='" + userName + "' and user_status=1";
            string s = DBHelper.GetValue(sql);
            if (s != "")
            {
                string userId = s.Split('/')[0];
                workCode = s.Split('/')[1];
                return userId;
            }
            workCode = "";
            return "";
        }

        /// <summary>
        /// 从指定目录中读取外地日考勤记录
        /// </summary>
        /// <returns></returns>
        static string ImportAttendanceFromExcel()
        {
            int recordCount = 0;
            string retMsg = "";
            DirectoryInfo di = new DirectoryInfo(attendanceDir);
            FileInfo[] arrFileInfo = di.GetFiles();
            foreach (FileInfo fi in arrFileInfo)
            {
                 string fileName = Path.GetFileName(fi.FullName);
                 string errMsg = "";
                 List<HrmScheduleSign> lstHrmScheduleSign= ImportExcel(fi,out errMsg);
                 if (lstHrmScheduleSign!=null && lstHrmScheduleSign.Count > 0)
                 {
                     recordCount = AddAttendance(lstHrmScheduleSign);
                     if (retMsg == "")
                     {
                         retMsg = "【"+fileName+"】:"+errMsg + "(同步考勤记录：" + recordCount + "条)";
                     }
                     else
                     {
                         retMsg += ";【" + fileName + "】:" + errMsg + "(同步考勤记录：" + recordCount + "条)";
                     }
                   
                 }
                 else
                 {
                     if (retMsg == "")
                     {
                         retMsg = errMsg;
                     }
                     else
                     {
                         retMsg += ";"+errMsg;
                     }
                 }
                 if (lstHrmScheduleSign!=null && lstHrmScheduleSign.Count > 0)
                 fi.Delete();
            }
            return retMsg;
        }

        /// <summary>
        /// 从Excel中读取日考勤记录
        /// </summary>
        /// <param name="fi">文件对象</param>
        /// <param name="errMsg">返回出错信息</param>
        static List<HrmScheduleSign> ImportExcel(FileInfo fi,out string errMsg)
        {
            List<HrmScheduleSign> lstHrmScheduleSign = new List<HrmScheduleSign>();
            IWorkbook workBook = null;
            Regex reg=null;
            Match match = null;
            string excelFileName=Path.GetFileName(fi.FullName);
            string excelType = ""; //分公司考勤记录类型
            string extName = ""; //文件扩展名
            string cellValue="";
            string tmpSignType = "";
            string month = ""; //月份
            int lastDay = 0; //每月最后天数
            int startRowIndex = 0; //开始行号
            foreach (string excelType2 in arrExcelType)
            {
                if (excelFileName.Contains(excelType2))
                {
                    excelType = excelType2;
                    break;
                }
            }
         
            FileStream streamExcel;
            try
            {
                 extName = Path.GetExtension(fi.FullName);
                 if (extName.ToLower() != ".xlsx" && extName.ToLower() != ".xls")
                 {
                     errMsg = "";
                     return null;
                 }
                 streamExcel = new FileStream(fi.FullName, FileMode.Open);
            }
            catch(Exception ex)
            {
                errMsg = excelType + "：" + ex.Message;
                return null;
            }
            //实例化工作薄
            extName = Path.GetExtension(fi.FullName);
            try
            {
                if (extName.ToLower() == ".xlsx")
                    workBook = new XSSFWorkbook(streamExcel);
                else
                    workBook = new HSSFWorkbook(streamExcel);
            }
            catch (Exception ex)
            {
                errMsg = excelType+"："+ ex.Message;
                return null;
            }
            ISheet workSheet = null;
           
            workSheet = workBook.GetSheetAt(0);
          
            if (workSheet == null)
            {
                errMsg = excelType+"：无效的工作表！";
                return null;
            } 
         
           
            int rowIndex = 0;
            IRow excelRow;
            string workCode = "";  //考勤机编号
            string userName = "";  //姓名
            string signDate = "";  //签到、签退日期
            string signTime = "";  //签到、时间
            string signType = "1"; //打卡类型(1:上班,2：下班)
            switch (excelType)
            {
                case "上海项目":
                    for (rowIndex = workSheet.FirstRowNum + 1; rowIndex <= workSheet.LastRowNum; rowIndex++)
                    {
                        excelRow = workSheet.GetRow(rowIndex);
                        workCode = "";  //考勤机编号
                        userName = "";  //姓名
                        signDate = "";  //签到、签退日期
                        signTime = "";  //签到、时间
                        signType = "1"; //打卡类型(1:上班,2：下班)
                        //考勤机编号
                        if (excelRow.GetCell(1) != null) workCode = excelRow.GetCell(1).StringCellValue + "(" + excelType+")";
                        //姓名
                        if (excelRow.GetCell(3) != null) userName = excelRow.GetCell(3).StringCellValue;
                        //签到日期
                        if (excelRow.GetCell(5) != null) signDate = excelRow.GetCell(5).StringCellValue;
                        //上班时间
                        if (excelRow.GetCell(9) != null) signTime = excelRow.GetCell(9).StringCellValue;
                        HrmScheduleSign hrmScheduleSign = new HrmScheduleSign();
                        hrmScheduleSign.WorkCode = workCode;
                        hrmScheduleSign.UserName = userName;
                        hrmScheduleSign.SignDate = !string.IsNullOrEmpty(signDate)?string.Format("{0:yyyy-MM-dd}",DateTime.Parse(signDate)):"";
                        hrmScheduleSign.SignTime = signTime;
                        hrmScheduleSign.SignType = signType;
                        lstHrmScheduleSign.Add(hrmScheduleSign);

                        //下班时间
                        if (excelRow.GetCell(10) != null) signTime = excelRow.GetCell(10).StringCellValue;
                        hrmScheduleSign = new HrmScheduleSign();
                        hrmScheduleSign.WorkCode = workCode;
                        hrmScheduleSign.UserName = userName;
                        hrmScheduleSign.SignDate = !string.IsNullOrEmpty(signDate) ? string.Format("{0:yyyy-MM-dd}", DateTime.Parse(signDate)) : "";
                        hrmScheduleSign.SignTime = signTime;
                        hrmScheduleSign.SignType = "2";
                        lstHrmScheduleSign.Add(hrmScheduleSign);
                    }
                   break;
                case "上海华沪":
                   for (rowIndex = workSheet.FirstRowNum + 1; rowIndex <= workSheet.LastRowNum; rowIndex++)
                   {
                       excelRow = workSheet.GetRow(rowIndex);
                       workCode = "";  //考勤机编号
                       userName = "";  //姓名
                       signDate = "";  //签到、签退日期
                       signTime = "";  //签到、时间
                       signType = "1"; //打卡类型(1:上班,2：下班)
                       //考勤机编号
                       if (excelRow!=null && excelRow.GetCell(0) != null)
                       {
                           switch (excelRow.GetCell(0).CellType)
                           {
                               case CellType.Numeric:
                                   workCode = excelRow.GetCell(0).NumericCellValue.ToString() + "(" + excelType + ")";
                                   break;
                               default:
                                   workCode = excelRow.GetCell(0).StringCellValue + "(" + excelType + ")";
                                   break;
                           }
                          
                       }
                       //姓名
                       if (excelRow!=null && excelRow.GetCell(2) != null) userName = excelRow.GetCell(2).StringCellValue;
                       //签到日期/时间
                       string signInfo = "";
                       if (excelRow!=null && excelRow.GetCell(3) != null)
                       {
                           switch (excelRow.GetCell(3).CellType)
                           {
                               case CellType.Numeric:
                                   signInfo = excelRow.GetCell(3).DateCellValue.ToString();
                                   break;
                               default:
                                   signInfo = excelRow.GetCell(3).StringCellValue;
                                   break;
                           }
                       }
                       if (signInfo != "")
                       {
                           string[] arrSignInfo = signInfo.Split(' ');
                           if (arrSignInfo.Length == 2)
                           {
                               signDate = arrSignInfo[0];
                               signTime = arrSignInfo[1] == "0:00:00" ? "" : arrSignInfo[1];
                           }
                           else
                           {
                               signDate = arrSignInfo[0];
                           }
                       }

                        HrmScheduleSign hrmScheduleSign = new HrmScheduleSign();
                        hrmScheduleSign.WorkCode = workCode;
                        hrmScheduleSign.UserName = userName;
                        hrmScheduleSign.SignDate = !string.IsNullOrEmpty(signDate) ? string.Format("{0:yyyy-MM-dd}", DateTime.Parse(signDate)) : "";
                        hrmScheduleSign.SignTime = signTime;
                        if (!string.IsNullOrEmpty(signTime))
                        {
                            if(DateTime.Parse(signTime).Hour<=12)
                               hrmScheduleSign.SignType = "1";
                            else
                               hrmScheduleSign.SignType = "2";
                            lstHrmScheduleSign.Add(hrmScheduleSign);

                        }
                        else
                        {
                            if (userName != "")
                            {
                                List<HrmScheduleSign> lstHrmScheduleSign2 = lstHrmScheduleSign.Where(p => (p.UserName == userName && p.SignDate == signDate)).ToList();
                                if (lstHrmScheduleSign2.Count == 0)
                                {
                                    hrmScheduleSign = new HrmScheduleSign();
                                    hrmScheduleSign.WorkCode = workCode;
                                    hrmScheduleSign.UserName = userName;
                                    hrmScheduleSign.SignDate = !string.IsNullOrEmpty(signDate) ? string.Format("{0:yyyy-MM-dd}", DateTime.Parse(signDate)) : "";
                                    hrmScheduleSign.SignType = "1";
                                }
                                else
                                {
                                    hrmScheduleSign = new HrmScheduleSign();
                                    hrmScheduleSign.WorkCode = workCode;
                                    hrmScheduleSign.UserName = userName;
                                    hrmScheduleSign.SignDate = !string.IsNullOrEmpty(signDate) ? string.Format("{0:yyyy-MM-dd}", DateTime.Parse(signDate)) : "";
                                    hrmScheduleSign.SignType = "2";
                                }
                                lstHrmScheduleSign.Add(hrmScheduleSign);
                            }
                        }
                      
                   }
                    break;
                case "济南项目":
                case "金奥":
                case "福中":
                    startRowIndex = 4;
                    int addRowNum = 6;
                    cellValue="";
                    string emptyFlag = "--:--";
                    excelRow = workSheet.GetRow(1);
                    cellValue = excelRow.GetCell(0).StringCellValue.Replace(" ", "");
                    string year = cellValue.Split(':').Length==2?cellValue.Split(':')[1]:DateTime.Now+"";
                    excelRow = workSheet.GetRow(2);
                    cellValue = excelRow.GetCell(0).StringCellValue.Replace(" ","");
                    month = cellValue.Split(':')[1];
                    signDate = year + "-" + month + "-";
                    lastDay = GetDaysOfMonth(month);
                    for (rowIndex = startRowIndex; rowIndex <= workSheet.LastRowNum; rowIndex+=addRowNum)
                    {
                        excelRow = workSheet.GetRow(rowIndex);
                        cellValue = excelRow.GetCell(1).StringCellValue;
                        if (cellValue != "")
                        {
                            string[] arrCell = cellValue.Split(new char[] {':',' '});
                            if (arrCell.Length > 0)
                            {
                                workCode = arrCell[2];
                                userName = arrCell[6];
                            }
                            for (int day = 1; day <= lastDay; day++)
                            {
                                HrmScheduleSign hrmScheduleSign = new HrmScheduleSign();
                                hrmScheduleSign.SignType = "1";
                                hrmScheduleSign.UserName = userName;
                                hrmScheduleSign.WorkCode = workCode;
                                hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                int signTimeRowIndex = rowIndex + 2;
                                if (day > 16)
                                {
                                      signTimeRowIndex = rowIndex + 4;
                                }
                                excelRow = workSheet.GetRow(signTimeRowIndex);
                                if(day<=16)
                                   cellValue = excelRow.GetCell(day).StringCellValue;
                                else
                                    cellValue = excelRow.GetCell(day-16).StringCellValue;
                                if (cellValue == "")
                                {
                                    lstHrmScheduleSign.Add(hrmScheduleSign);
                                    hrmScheduleSign = new HrmScheduleSign();
                                    hrmScheduleSign.SignType = "2";
                                    hrmScheduleSign.UserName = userName;
                                    hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                    hrmScheduleSign.WorkCode = workCode;
                                    lstHrmScheduleSign.Add(hrmScheduleSign);
                                }
                                else
                                {
                                    string[] arrSignTime = cellValue.Split(' ');
                                    hrmScheduleSign.SignTime = arrSignTime[0] != emptyFlag ? arrSignTime[0] : "";
                                    if (DateTime.Parse(hrmScheduleSign.SignTime).Hour <= 12)
                                        hrmScheduleSign.SignType = "1";
                                    else
                                        hrmScheduleSign.SignType = "2";
                                    tmpSignType= hrmScheduleSign.SignType;
                                    lstHrmScheduleSign.Add(hrmScheduleSign);
                                    hrmScheduleSign = new HrmScheduleSign();
                                    hrmScheduleSign.SignType = tmpSignType=="1"?"2":"1";
                                    hrmScheduleSign.UserName = userName;
                                    hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                    hrmScheduleSign.WorkCode = workCode;
                                 
                                    if (arrSignTime.Length > 2)
                                        hrmScheduleSign.SignTime = arrSignTime[arrSignTime.Length - 2];
                                    else
                                        hrmScheduleSign.SignTime = arrSignTime[arrSignTime.Length - 1];
                                    
                                    lstHrmScheduleSign.Add(hrmScheduleSign);
                                }
                            }
                        }
                    }
                    break;
                case "苏州项目":
                    for (rowIndex = workSheet.FirstRowNum + 1; rowIndex <= workSheet.LastRowNum; rowIndex++)
                    {
                        excelRow = workSheet.GetRow(rowIndex);
                        workCode = "";  //考勤机编号
                        userName = "";  //姓名
                        signDate = "";  //签到、签退日期
                        signTime = "";  //签到、时间
                        signType = "1"; //打卡类型(1:上班,2：下班)
                        //姓名
                        if (excelRow.GetCell(0) != null) userName = excelRow.GetCell(0).StringCellValue;
                        //签到日期/时间
                        string signInfo = "";
                        if (excelRow.GetCell(1) != null)
                        {
                            switch (excelRow.GetCell(1).CellType)
                            {
                                case CellType.Numeric:
                                    signInfo = excelRow.GetCell(1).DateCellValue.ToString();
                                    break;
                                default:
                                    signInfo = excelRow.GetCell(1).StringCellValue;
                                    break;
                            }
                        }
                        if (signInfo != "")
                        {
                            string[] arrSignInfo = signInfo.Split(' ');
                            if (arrSignInfo.Length == 2)
                            {
                                signDate = arrSignInfo[0];
                                signTime = arrSignInfo[1] == "0:00:00" ? "" : arrSignInfo[1];
                            }
                            else
                            {
                                signDate = arrSignInfo[0];
                            }
                        }

                        HrmScheduleSign hrmScheduleSign = new HrmScheduleSign();
                        hrmScheduleSign.WorkCode = workCode;
                        hrmScheduleSign.UserName = userName;
                        hrmScheduleSign.SignDate = !string.IsNullOrEmpty(signDate) ? string.Format("{0:yyyy-MM-dd}", DateTime.Parse(signDate)) : "";
                        hrmScheduleSign.SignTime = signTime;
                        if (!string.IsNullOrEmpty(signTime))
                        {
                            if (DateTime.Parse(signTime).Hour <=12)
                                hrmScheduleSign.SignType = "1";
                            else
                                hrmScheduleSign.SignType = "2";
                            lstHrmScheduleSign.Add(hrmScheduleSign);

                        }
                        else
                        {
                            List<HrmScheduleSign> lstHrmScheduleSign2 = lstHrmScheduleSign.Where(p => (p.UserName == userName && p.SignDate == signDate)).ToList();
                            if (lstHrmScheduleSign2.Count == 0)
                            {
                                hrmScheduleSign.SignType = "1";
                            }
                            else
                            {
                                hrmScheduleSign.SignType = "2";
                            }
                            lstHrmScheduleSign.Add(hrmScheduleSign);
                        }

                    }
                    break;
                case "金润":
                    startRowIndex = 4;
                    excelRow = workSheet.GetRow(2);
                    cellValue = excelRow.GetCell(2).StringCellValue.Replace(" ","");
                    reg = new Regex(@"\d{4}/\d{2}/\d{2}");
                    match= reg.Match(cellValue);
                    if (match.Success)
                    {
                        month = DateTime.Parse(match.Value).Month + "";
                        signDate = DateTime.Parse(match.Value).Year + "-" + DateTime.Parse(match.Value).Month.ToString("00")+ "-";
                        lastDay = GetDaysOfMonth(month);
                        for (rowIndex = startRowIndex; rowIndex <= workSheet.LastRowNum; rowIndex+=2)
                        {
                            excelRow = workSheet.GetRow(rowIndex);
                            cellValue = excelRow.GetCell(2).StringCellValue;
                            if (cellValue != "")
                            {
                                workCode = cellValue + "(" + excelType + ")";
                            }
                            cellValue = excelRow.GetCell(10).StringCellValue;
                            if (cellValue != "")
                            {
                                userName = cellValue;
                            }
                            excelRow = workSheet.GetRow(rowIndex+1);
                            for (int day = 1; day <= lastDay; day++)
                            {
                                HrmScheduleSign hrmScheduleSign = new HrmScheduleSign();
                                hrmScheduleSign.WorkCode = workCode;
                                hrmScheduleSign.UserName = userName;
                                cellValue = excelRow.GetCell(day - 1).StringCellValue;
                                if (cellValue != "")
                                {
                                  
                                    string[] arrSignTime = cellValue.Replace(" ","").Split('\n');
                                    if (arrSignTime.Length == 3)
                                    {
                                        hrmScheduleSign.SignType = "1";
                                        hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                        signTime = arrSignTime[0];
                                        hrmScheduleSign.SignTime = signTime;
                                        lstHrmScheduleSign.Add(hrmScheduleSign);

                                        hrmScheduleSign = new HrmScheduleSign();
                                        hrmScheduleSign.WorkCode = workCode;
                                        hrmScheduleSign.UserName = userName;
                                        hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                        hrmScheduleSign.SignType = "2";
                                        signTime = arrSignTime[1];
                                        hrmScheduleSign.SignTime = signTime;
                                        lstHrmScheduleSign.Add(hrmScheduleSign);
                                    }
                                    else
                                    {
                                        DateTime dtSignTime = DateTime.Parse(arrSignTime[0]);
                                        if (dtSignTime.Hour <= 12)
                                        {
                                            hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                            hrmScheduleSign.SignType = "1";
                                            signTime = arrSignTime[0];
                                            hrmScheduleSign.SignTime = signTime;
                                            lstHrmScheduleSign.Add(hrmScheduleSign);

                                            hrmScheduleSign = new HrmScheduleSign();
                                            hrmScheduleSign.WorkCode = workCode;
                                            hrmScheduleSign.UserName = userName;
                                            hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                            hrmScheduleSign.SignType = "2";
                                            lstHrmScheduleSign.Add(hrmScheduleSign);
                                        }
                                        else
                                        {
                                            hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                            hrmScheduleSign.SignType = "2";
                                            signTime = arrSignTime[0];
                                            hrmScheduleSign.SignTime = signTime;
                                            lstHrmScheduleSign.Add(hrmScheduleSign);

                                            hrmScheduleSign = new HrmScheduleSign();
                                            hrmScheduleSign.WorkCode = workCode;
                                            hrmScheduleSign.UserName = userName;
                                            hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                            hrmScheduleSign.SignType = "1";
                                            lstHrmScheduleSign.Add(hrmScheduleSign);
                                        }
                                    }
                                }
                                else
                                {
                                    hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                    hrmScheduleSign.SignType = "1";
                                    lstHrmScheduleSign.Add(hrmScheduleSign);
                                    hrmScheduleSign = new HrmScheduleSign();
                                    hrmScheduleSign.WorkCode = workCode;
                                    hrmScheduleSign.UserName = userName;
                                    hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                    hrmScheduleSign.SignType = "2";
                                    lstHrmScheduleSign.Add(hrmScheduleSign);
                                }
                            }
                        }
                    }
                    else
                    {
                        errMsg = excelType+"：无效的日期格式！";
                        return lstHrmScheduleSign;
                    }
                    break;
                case "商管":
                    excelRow = workSheet.GetRow(2);
                    cellValue = excelRow.GetCell(2).StringCellValue.Replace(" ","");
                    reg = new Regex(@"\d{4}-\d{2}-\d{2}");
                    match= reg.Match(cellValue);
                    if (match.Success)
                    {
                        month = DateTime.Parse(match.Value).Month + "";
                        signDate = DateTime.Parse(match.Value).Year + "-" + DateTime.Parse(match.Value).Month.ToString("00") + "-";
                    }
                    startRowIndex = 4;
                    reg = new Regex(@"\d{2}:\d{2}");
                    for (rowIndex = startRowIndex; rowIndex <= workSheet.LastRowNum; rowIndex += 2)
                    {
                         excelRow = workSheet.GetRow(rowIndex);
                         try
                         {
                             workCode = excelRow.GetCell(2).StringCellValue.Replace(" ", "");
                         }
                         catch
                         {
                             workCode = excelRow.GetCell(2).NumericCellValue.ToString();
                         }
                         workCode = workCode + "(" + excelType + ")";
                         if (excelRow.GetCell(10) == null)
                         {
                             continue;
                         }
                         userName = excelRow.GetCell(10).StringCellValue.Replace(" ", "");
                         excelRow = workSheet.GetRow(3);
                         int cellCount = excelRow.LastCellNum;
                         for (int day = 1; day <= cellCount; day++)
                         {
                             HrmScheduleSign hrmScheduleSign = new HrmScheduleSign();
                             hrmScheduleSign.WorkCode = workCode;
                             hrmScheduleSign.UserName = userName;
                             hrmScheduleSign.SignDate = signDate + day.ToString("00");
                             excelRow = workSheet.GetRow(rowIndex + 1);
                             string signTimeInfo = (excelRow != null && excelRow.GetCell(day - 1)!=null) ? excelRow.GetCell(day - 1).StringCellValue : "";
                             match = reg.Match(signTimeInfo);
                             if (match.Success)
                             {
                                 hrmScheduleSign.SignTime = match.Value;
                                 if(DateTime.Parse( hrmScheduleSign.SignTime).Hour<12) 
                                     hrmScheduleSign.SignType = "1";
                                 else
                                     hrmScheduleSign.SignType = "2";
                                 lstHrmScheduleSign.Add(hrmScheduleSign);

                                 hrmScheduleSign = new HrmScheduleSign();
                                 hrmScheduleSign.WorkCode = workCode;
                                 hrmScheduleSign.UserName = userName;
                                 hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                 if (match.NextMatch().Success)
                                 {
                                     hrmScheduleSign.SignTime = match.NextMatch().Value;
                                 }
                                 hrmScheduleSign.SignType = tmpSignType == "1" ? "2" : "1";
                                 lstHrmScheduleSign.Add(hrmScheduleSign);

                             }
                             else
                             {
                                 hrmScheduleSign.SignType = "1";
                                 lstHrmScheduleSign.Add(hrmScheduleSign);

                                 hrmScheduleSign = new HrmScheduleSign();
                                 hrmScheduleSign.WorkCode = workCode;
                                 hrmScheduleSign.UserName = userName;
                                 hrmScheduleSign.SignDate = signDate + day.ToString("00");
                                 hrmScheduleSign.SignType = "2";
                                 lstHrmScheduleSign.Add(hrmScheduleSign);
                             }
                         }
                    }
                    break;
            }
            streamExcel.Close();
            errMsg = excelType;
            return lstHrmScheduleSign;
        }

        /// <summary>
        /// 获取指定月份的天数
        /// </summary>
        /// <param name="month">指定月份</param>
        /// <returns></returns>
        static int GetDaysOfMonth(string month)
        {
            DateTime dt = new DateTime(DateTime.Now.Year, int.Parse(month), 1);
            dt= dt.AddDays(1 - dt.Day).AddMonths(1).AddDays(-1);
            return dt.Day;
        }
    }
}
