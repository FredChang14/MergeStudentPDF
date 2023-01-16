using Google.Protobuf.WellKnownTypes;
using iText.Layout.Element;
using MergeStudentPDF2.Librarys;
using MergeStudentPDF2.Model;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace MergeStudentPDF2
{
    internal class Program
    {

        static int Success = 0;

        static int Fail = 0;

        static void Main(string[] args)
        {
            JObject Setting = JObject.Parse(File.ReadAllText("Settings.json"));

            if (Setting["學生data所在路徑"].ToString() != string.Empty)
            {
                if (!Directory.Exists(Setting["學生data所在路徑"].ToString()))
                {
                    Console.WriteLine("學生data資料夾所在路徑找不到，請確認Settings.json 路徑是否正確");

                    Console.ReadKey();

                    return;
                } // end if

                StudentPdfHandle.SetDataParnetDirect(Setting["學生data所在路徑"].ToString());
            } // end if

            DBHelper.SetConnectionInfo(Setting["資料庫IP"].ToString(), Setting["資料庫連接埠"].ToString(), Setting["資料庫名稱"].ToString(), Setting["DBUser"].ToString(), Setting["Pwd"].ToString());

            if (!DBHelper.TestConnection())
            {
                Console.WriteLine("資料庫連線測試失敗，請確認Settings.json資料庫連線資訊是否正確");

                Console.ReadKey();

                return;
            } // end if

            if (!Directory.Exists("Merge"))
            {
                Directory.CreateDirectory("Merge");
            } // end if
            //TestMergePDF1();
            MergePDF();

            Console.ReadKey();
        }


        static void MergePDF()
        {
            System.Collections.Generic.List<StudentInfo> StudentSIDList = EPItemTask.GetStudentSIDList();

            Stopwatch SW = new Stopwatch();

            SW.Reset();

            SW.Start();

            for (int i = 0; i < StudentSIDList.Count; i++)
            {
                int SID = StudentSIDList[i].SID;

                string DepartmentCode = StudentSIDList[i].DepartmentCode;

                string RegNum = StudentSIDList[i].Reg_Num;

                bool Result = Excute(SID, DepartmentCode, RegNum);

                if (Result)
                {
                    Console.WriteLine(string.Format("{0}-{1} 學生合併PDF 產生成功", DepartmentCode, RegNum));
                }
                else
                {
                    Console.WriteLine(string.Format("{0}-{1} 學生合併PDF 產生失敗", DepartmentCode, RegNum));
                } // end if
            } // end for

            SW.Stop();

            Console.WriteLine("==============================================");
            Console.WriteLine("執行成功");
            Console.WriteLine("總筆數 {0}", StudentSIDList.Count);
            Console.WriteLine("成功筆數 {0}", Success);
            Console.WriteLine("失敗筆數 {0} ", Fail);
            Console.WriteLine("花費時間 : {0}", SW.Elapsed);
            string ErrMsg = string.Format($"=============================================={Environment.NewLine}執行程式時間: {DateTime.UtcNow.AddHours(8).ToString("yyyy-MM-dd HH:mm:ss")}{Environment.NewLine}總筆數 : {StudentSIDList.Count}{Environment.NewLine}成功筆數 : {Success}{Environment.NewLine}失敗筆數 : {Fail}{Environment.NewLine}花費時間 : {SW.Elapsed}{Environment.NewLine}=============================================={Environment.NewLine}");
            LogTask.WriteLogMessage(ErrMsg);
        } // end MergePDF


        // 單筆測試
        static void TestMergePDF1()
        {
            Excute(6669, "000001", "38419220");

            Console.WriteLine("Success");
        } // end TestMergePDF


        // 多執行續執行
        static void TestMergePDF2()
        {
            Stopwatch SW = new Stopwatch();

            SW.Reset();

            SW.Start();

            System.Collections.Generic.List<StudentInfo> StudentSIDList = EPItemTask.GetStudentSIDList();

            Task[] tasks = new Task[StudentSIDList.Count];

            for (int i = 0; i < StudentSIDList.Count; i++)
            {
                int SID = StudentSIDList[i].SID;

                string DepartmentCode = StudentSIDList[i].DepartmentCode;

                string RegNum = StudentSIDList[i].Reg_Num;

                tasks[i] = Task.Factory.StartNew(() => Excute(SID, DepartmentCode, RegNum));
            } // end for

            Task.Factory.ContinueWhenAll(tasks, completedTasks =>
            {
                SW.Stop();
                Console.WriteLine("==============================================");
                Console.WriteLine("執行成功");
                Console.WriteLine("總筆數 {0}", StudentSIDList.Count);
                Console.WriteLine("成功筆數 {0}", Success);
                Console.WriteLine("失敗筆數 {0} ", Fail);
                Console.WriteLine("花費時間 : {0}", SW.Elapsed);
                string ErrMsg = string.Format("==============================================\n執行程式時間: {0}\n總筆數 : {1}\n成功筆數 : {2}\n失敗筆數 : {3}\n花費時間 : {4}\n==============================================\n",  DateTime.UtcNow.AddHours(8).ToString("yyyy-MM-dd HH:mm:ss"), StudentSIDList.Count, Success, Fail, SW.Elapsed);
                LogTask.WriteLogMessage(ErrMsg);
            });
        } // end TestMergePDF2


        static bool Excute(int SID, string DepartmentCode, string RegNum)
        {
            bool Result = true;

            try
            {
                StudentPdfHandle.StartMerge(SID);

                Success++;

                Result = true;
            }
            catch (Exception ex)
            {
                Fail++;

                string ErrMsg = string.Format("{0}-{1}  PDF合併失敗(可能有檔案毀損或路徑不正確) : {2}", DepartmentCode, RegNum, ex.Message);

                LogTask.WriteLogMessage(ErrMsg);

                Result = false;
            } // end try catch

            return Result;
        }
    }
}
