using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MergeStudentPDF2.Librarys
{
    public class LogTask
    {

        private static object LockObj = new object();

        public static void WriteLogMessage(string ErrMsg)
        {
            lock (LockObj)
            {
                File.AppendAllText("Log.txt", ErrMsg + Environment.NewLine);
            } // end lock
        } // end WriteLogMessage
    } // end LogTask
}
