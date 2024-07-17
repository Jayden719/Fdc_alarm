using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using System.Data.SqlClient;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;

namespace Fdc_alarm
{
    public partial class Form1 : Form
    {
        string LogPath = Application.StartupPath + "\\log\\log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
        string FolderPath = Application.StartupPath + "\\log";      


        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileInt(string section, string key, int val,  string filepath);       
        string iniDrectory = Application.StartupPath + @"\fdc_val.ini";
             
        public int ReadINI(string sec, string keyname, string fp)
        {
            int time;
            time = GetPrivateProfileInt(sec, keyname, 0, fp);
            return time;
        }

        // 타이머 객체 생성
        System.Timers.Timer timer_start = new System.Timers.Timer();
        System.Timers.Timer timer1 = new System.Timers.Timer();
        System.Timers.Timer timer2 = new System.Timers.Timer();

        int hours = 0;
        int interval = 0;
        int start_h = 0;
        int start_m = 0;
        int end_h = 0;
        int end_m = 0;
        bool send_log = false;
        bool board_log = false;

        // 파일 스트림 사용여부 판단 메소드
        public static bool IsFileUseGeneric(FileInfo file)
        {
            try
            {
                using var stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            return false;
        }


        private void LogWriter()
        {
            // 로그 폴더(디렉토리) 존재 여부 판단하여 생성
            DirectoryInfo di = new DirectoryInfo(FolderPath);
            if (di.Exists == true)
            {

            }
            else
            {
                di.Create();
            }

            // 당일 로그 파일 존재 여부 판단하여 생성하기
            string LogPath_n = Application.StartupPath + "\\log\\log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            if (LogPath == LogPath_n)
            {
                // 존재여부
                System.IO.FileInfo fi = new System.IO.FileInfo(LogPath);
                if (fi.Exists == true)
                {

                }
                else
                {
                    Console.WriteLine("프로그램 수동 실행 시 여기로 온다");
                   /* FileStream file = File.Open(LogPath_n, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    StreamWriter sw = new StreamWriter(file);
                    sw.Write(DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " 로그 파일 최초 생성 ");*/
                    System.IO.File.AppendAllText(LogPath_n, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " 로그 파일 최초 생성 ", Encoding.Default);
                }
                // 객제 해제를 위한 null 입력 → GC
                fi = null;
            }
            else
            {
                System.IO.File.AppendAllText(LogPath_n, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " 로그 파일 최초 생성 ", Encoding.Default);
                LogPath = LogPath_n; // 전역변수에 최신화된 데이터 저장
                send_log = true;
                board_log = true;
            }
           

           /* LogPath 전역변수 값이 유지되어 날짜 지나도 로그파일 생성 못함
            * System.IO.FileInfo fi = new System.IO.FileInfo(LogPath);
            if (fi.Exists == true)
            {

            }
            else
            {
                System.IO.File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " 로그 파일 최초 생성 ", Encoding.Default);
            }*/
        }

        public Form1()
        {
            InitializeComponent();
            Console.WriteLine("팩스 서버 알림 프로그램 실행중...");
            this.Text = string.Format("팩스 알림 프로그램 ver_{0}", Application.ProductVersion);
            this.ShowInTaskbar = false;
            this.Opacity = 0;

            start_h = ReadINI("Fax_Timer", "start_hour", iniDrectory);
            start_m = ReadINI("Fax_Timer", "start_min", iniDrectory);
            end_h = ReadINI("Fax_Timer", "end_hour", iniDrectory);
            end_m = ReadINI("Fax_Timer", "end_min", iniDrectory);
          
            if(start_h < 10)
            {
                start_h = Convert.ToInt32("0" + start_h.ToString());
            }
            if(start_m < 10)
            {
                start_m = Convert.ToInt32("0" + start_m.ToString());
            }
            if (end_h < 10)
            {
                end_h = Convert.ToInt32("0" + end_h.ToString());
            }
            if (end_m < 10)
            {
                end_m = Convert.ToInt32("0" + end_m.ToString());
            }

            if(Convert.ToInt32(DateTime.Now.ToString("HHmm")) < start_h * 100 + start_m || Convert.ToInt32(DateTime.Now.ToString("HHmm")) > end_h * 100 + end_m)
            {
                MessageBox.Show("시간 설정을 다시하세요");
                return;
            }
            else
            {
                timer_start.Elapsed += new ElapsedEventHandler(Fdc_alarm);
                timer_start.Start();
                LogWriter();

            }    
        }


        // 프로그램 첫 실행 메소드
        private void Fdc_alarm(object sender, ElapsedEventArgs e)
        {                                              
            LogWriter();
            string LogPath_n = Application.StartupPath + "\\log\\log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Fdc_alarm 타이머 동작", Encoding.Default);               
          
            // 종료시간을 초과한 경우
            int endH = ReadINI("Fax_Timer", "end_hour", iniDrectory);
            int endM = ReadINI("Fax_Timer", "end_min", iniDrectory);
            if(endH < 10)
            {
                endH = Convert.ToInt32("0" + endH.ToString());
            }
            if(endM < 10)
            {
                endM = Convert.ToInt32("0" + endM.ToString());
            }
            int now = Convert.ToInt32(DateTime.Now.ToString("HHmm"));
            int now_h = Convert.ToInt32(DateTime.Now.ToString("HH"));
            int now_m= Convert.ToInt32(DateTime.Now.ToString("mm"));           

            // 당일 시작 시간에 맞춰서 프로그램 실행된 경우
            if(now < start_h*100+start_m + 1 && start_h * 100 + start_m <= now)
            {
                // 여기서 24시간 후에 다시 반복 실행
                hours = 24;
                timer_start.Interval = hours * 60 * 60 * 1000;
            }
            else
            {
                // 09:00 ~ 16:00 사이에 실행되어 다음날 시작시간부터 시작
                interval = ((23 - now_h + start_h) * 60 * 60 + (60 - now_m) * 60) * 1000;
                timer_start.Interval = interval;
            }
  

            // 팩스 발송건 조회                  
            
            if (send_log)
            {
                send_log = false;
                send_alarm();
                timer1.Elapsed += new ElapsedEventHandler(Send_alarm_test);
                timer1.Start();
            }
            else
            {
                timer1.Elapsed += new ElapsedEventHandler(Send_alarm_test);
                timer1.Start();
            }

            // 보드 성공률 조회                
            if (board_log)
            {
                board_log = false;
                board_alarm();
                timer2.Elapsed += new ElapsedEventHandler(Board_alarm_timer);
                timer2.Start();
            }
            else
            {
                timer2.Elapsed += new ElapsedEventHandler(Board_alarm_timer);
                timer2.Start();
            }

        }

        private void Board_alarm_timer(object sender, ElapsedEventArgs e)
        {
            timer2.Interval = 60 * 60 * 1000; //1시간 단위로 반복
            board_alarm();          
        }

        private void board_alarm()
        {
            string LogPath_n = Application.StartupPath + "\\log\\log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            if(IsFileUseGeneric(new FileInfo(LogPath_n)))
            {
                Thread.Sleep(5000);
                File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 타이머 시작", Encoding.Default);
            }

            int endH = ReadINI("Fax_Timer", "end_hour", iniDrectory);
            int endM = ReadINI("Fax_Timer", "end_min", iniDrectory);

            string db_sql = "Server=222.231.58.75; database=FAX; uid=eshinan; pwd=!eshinan4600";
            List<string> boardList = new List<string>();
            string srvId = "";
            string ntot = "";
            string nsuc = "";
            string nfai = "";
            string srate = "";
            string board_res = "";
            string board_alarm = "";

            if (Convert.ToInt32(DateTime.Now.ToString("HHmm")) >= endH * 100 + endM)
            {
                if (IsFileUseGeneric(new FileInfo(LogPath_n)))
                {
                    Thread.Sleep(5000);
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 타이머 종료", Encoding.Default);
                    //timer_cnt += 1;
                    timer2.Enabled = false;
                    timer2.Stop();
                    /*if(timer_cnt == 2)
                    {
                        timer_start.Stop();
                        File.AppendAllText(LogPath, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Fdc_alarm 타이머 종료", Encoding.Default);
                    }*/
                }
                else
                {
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 타이머 종료", Encoding.Default);
                    timer2.Enabled = false;
                    timer2.Stop();
                }
                return;
            }

            using (SqlConnection conn = new SqlConnection(db_sql))
            {
                conn.Open();
                string startDate = DateTime.Now.ToString("yyyy-MM-dd");
                string endDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");

                string fdc_sql = "select a.sSrvID as 서버아이디, a.전체건 as 전체건수, b.성공건, (a.전체건 - b.성공건) as 실패건, ROUND(CONVERT(float, ISNULL(b.성공건,0 )) / a.전체건, 2) *100 as 성공률 " +
                    "from (select sSrvID, count(*) as 전체건 from SentLog(nolock) where dtStartTime between DATEADD(MINUTE, -60, GETDATE()) " +
                    "and  Getdate()  group by sSrvID) a left outer join (select sSrvID, count(*) as 성공건 from " +
                    "SentLog(nolock) where dtStartTime between DATEADD(MINUTE, -60, GETDATE()) and Getdate() and nResult in (0, 777, 778, 773, 770, 779)  group by sSrvID) b on a.sSrvID = b.sSrvID " +
                    "where ROUND(CONVERT(float, ISNULL(b.성공건,0 )) / a.전체건, 2) * 100 < 75";

                SqlCommand comm = new SqlCommand(fdc_sql, conn);
                SqlDataReader rdr = comm.ExecuteReader();

                while (rdr.Read())
                {
                    srvId = rdr["서버아이디"].ToString().Replace(" ", "");
                    if (srvId == "1")
                    {
                        srvId = "222.231.58.101";
                    }
                    else if (srvId == "2")
                    {
                        srvId = "222.231.58.102";
                    }
                    else if (srvId == "4")
                    {
                        srvId = "222.231.58.104";
                    }
                    else if (srvId == "5")
                    {

                        srvId = "222.231.58.109";
                    }
                    else if (srvId == "6")
                    {
                        srvId = "222.231.58.117";
                    }
                    else if (srvId == "7")
                    {
                        srvId = "222.231.58.107";
                    }
                    else if (srvId == "8")
                    {
                        srvId = "222.231.58.103";
                    }
                    else if (srvId == "9")
                    {
                        srvId = "222.231.58.105";
                    }
                    else if (srvId == "10")
                    {
                        srvId = "222.231.58.110";
                    }
                    else if (srvId == "11")
                    {
                        srvId = "222.231.58.111";
                    }
                    else if (srvId == "12")
                    {
                        srvId = "222.231.58.112";
                    }
                    else if (srvId == "13")
                    {
                        srvId = "222.231.58.72";
                    }
                    else if (srvId == "14")
                    {
                        srvId = "222.231.58.118";
                    }
                    else if (srvId == "15")
                    {
                        srvId = "222.231.58.81";
                    }
                    else if (srvId == "16")
                    {
                        srvId = "222.231.58.83";
                    }
                    else if (srvId == "17")
                    {
                        srvId = "222.231.58.108";
                    }
                    else if (srvId == "18")
                    {
                        srvId = "222.231.58.123";
                    }
                    ntot = string.Format("{0:#,###}", rdr["전체건수"]);
                    nsuc = string.Format("{0:#,###}", rdr["성공건"]);
                    nfai = string.Format("{0:#,###}", rdr["실패건"]);
                    srate = rdr["성공률"].ToString() + "%";
                    board_res = string.Format("[보드 장애 알림 ({0})기준 ]\r\n{1} (성공률 : {2})\r\n■ 전체건수 : {3}건\r\n■ 성공건수 : {4}건\r\n■ 실패건수 : {5}건\r\n", DateTime.Now.ToString("yy-MM-dd hh:mm"), srvId, srate, ntot, nsuc, nfai);
                    boardList.Add(board_res);
                }
                rdr.Close();
                comm.Dispose();
                conn.Close();
            }
            if (boardList.Count == 0)
            {
                if (IsFileUseGeneric(new FileInfo(LogPath_n)))
                {
                    Thread.Sleep(5000);
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 알림 이상 없습니다", Encoding.Default);
                }
                else
                {
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 알림 이상 없습니다", Encoding.Default);

                }
            }
            else
            {
                foreach (string b in boardList)
                {
                    board_alarm += b + "\n";
                }
                if (IsFileUseGeneric(new FileInfo(LogPath_n)))
                {
                    Thread.Sleep(5000);
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 알림 문자 발송", Encoding.Default);
                }
                else
                {
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Board_alarm 알림 문자 발송", Encoding.Default);

                }
                MessageBox.Show(board_alarm);
            }
        }

        private void Send_alarm_test(object sender, ElapsedEventArgs e)
        {
            timer1.Interval = 60 * 60 * 1000; // 1시간 단위로 반복
            send_alarm();
            
        }
        private void send_alarm()
        {
            string LogPath_n = Application.StartupPath + "\\log\\log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            if(IsFileUseGeneric(new FileInfo(LogPath_n)))
            {
                Thread.Sleep(5000);
                File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 타이머 시작", Encoding.Default);
            }

            int endH = ReadINI("Fax_Timer", "end_hour", iniDrectory);
            int endM = ReadINI("Fax_Timer", "end_min", iniDrectory);
            int pageTime = ReadINI("Fax_Timer", "page_time", iniDrectory);

            List<string> sendList = new List<string>();
            string send_alarm = "";
            string srv = "";
            string smod = "";
            string sjob = "";
            string dts = "";
            string ntot = "";
            string send_res = "";

            if (Convert.ToInt32(DateTime.Now.ToString("HHmm")) >= endH * 100 + endM)
            {
                // timer_cnt += 1;
                if (IsFileUseGeneric(new FileInfo(LogPath_n)))
                {
                    Thread.Sleep(5000);
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 타이머 종료", Encoding.Default);
                }
                else
                {
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 타이머 종료", Encoding.Default);

                }
                timer1.Enabled = false;
                timer1.Stop();
                /*if (timer_cnt == 2)
                {
                    timer_start.Stop();
                    File.AppendAllText(LogPath, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Fdc_alarm 타이머 종료", Encoding.Default);
                }*/
                return;
            }

            string db_sql = "Server=222.231.58.75; database=FAX; uid=eshinan; pwd=!eshinan4600";
            using (SqlConnection conn = new SqlConnection(db_sql))
            {
                conn.Open();
                string fdc_sql = string.Format("select a.sSrvid as 서버명, a.sModemID as 보드, b.sjobid as 잡아이디, b.dtstarttime as 전송시간, b.ntotalpage as 전체페이지 from SentLog(nolock) a \r\n" +
                    "inner join TaskFaxLog(nolock) b\r\non a.sJobID = b.sJobID\r\nwhere b.sJobID in \r\n(select sJobID from TaskFaxLog(nolock) where sJobId in " +
                    "(\r\nselect sjobid from JobfaxLog(nolock) where ROUND(convert(float, nFileSize) / nTotalPage,2) < 100 and nTotalPage != 0 and nTotalPage is not null " +
                    "and dtStartTime > DATEADD(DAY, -2, GETDATE())\r\n) " +
                    "group by sJobID having count(sJobID) = 1)\r\n" +
                    "and (b.nSentSrv=1 and b.sSentSrvID is not null and b.sSentSrvID !='') and b.dtEndTime is null and b.nCancel = 0\r\n" +
                    "and DATEADD(MINUTE, {0}*b.nTotalPage ,b.dtStartTime) < GETDATE()", pageTime);

                SqlCommand comm = new SqlCommand(fdc_sql, conn);
                SqlDataReader rdr = comm.ExecuteReader();
                while (rdr.Read())
                {
                    srv = rdr["서버명"].ToString().Replace(" ", "");
                    if (srv == "1")
                    {
                        srv = "222.231.58.101";
                    }
                    else if (srv == "2")
                    {
                        srv = "222.231.58.102";
                    }
                    else if (srv == "4")
                    {
                        srv = "222.231.58.104";
                    }
                    else if (srv == "5")
                    {
                        srv = "222.231.58.109";
                    }
                    else if (srv == "6")
                    {
                        srv = "222.231.58.117";
                    }
                    else if (srv == "7")
                    {
                        srv = "222.231.58.107";
                    }
                    else if (srv == "8")
                    {
                        srv = "222.231.58.103";
                    }
                    else if (srv == "9")
                    {
                        srv = "222.231.58.105";
                    }
                    else if (srv == "10")
                    {
                        srv = "222.231.58.110";
                    }
                    else if (srv == "11")
                    {
                        srv = "222.231.58.111";
                    }
                    else if (srv == "12")
                    {
                        srv = "222.231.58.112";
                    }
                    else if (srv == "13")
                    {
                        srv = "222.231.58.72";
                    }
                    else if (srv == "14")
                    {
                        srv = "222.231.58.118";
                    }
                    else if (srv == "15")
                    {
                        srv = "222.231.58.81";
                    }
                    else if (srv == "16")
                    {
                        srv = "222.231.58.83";
                    }
                    else if (srv == "17")
                    {
                        srv = "222.231.58.108";
                    }
                    else if (srv == "18")
                    {
                        srv = "222.231.58.123";
                    }
                    smod = rdr["보드"].ToString();
                    sjob = rdr["잡아이디"].ToString();
                    dts = rdr["전송시간"].ToString();
                    ntot = string.Format("{0:#,###}", rdr["전체페이지"]);

                    send_res = string.Format("[FDC발송 장애 알림 ({0})기준]\r\n{1}\r\n■ 전송시간 : {2}\r\n■ 잡아이디 : {3}\r\n■ 전체페이지 : {4}장\r\n", DateTime.Now.ToString("yy-MM-dd HH:mm"), srv, dts, sjob, ntot);
                    sendList.Add(send_res);
                }
                rdr.Close();
                comm.Dispose();
                conn.Close();
            }
            if (sendList.Count == 0)
            {
                /*FileStream file = File.Open(LogPath_n, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                StreamWriter sw = new StreamWriter(file);
                sw.Write("\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 알림 이상없습니다");*/
                if (IsFileUseGeneric(new FileInfo(LogPath_n)))
                {
                    Thread.Sleep(5000);
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 알림 이상없습니다", Encoding.Default);
                }
                else
                {
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 알림 이상없습니다", Encoding.Default);

                }
            }
            else
            {
                foreach (string s in sendList)
                {
                    send_alarm += s + "\n";
                }
                if (IsFileUseGeneric(new FileInfo(LogPath_n)))
                {
                    Thread.Sleep(5000);
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 알림 문자 발송", Encoding.Default);

                }
                else
                {
                    File.AppendAllText(LogPath_n, "\r\n" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " Send_alarm 알림 문자 발송", Encoding.Default);

                }
                MessageBox.Show(send_alarm);
            }
        }
    }
}
