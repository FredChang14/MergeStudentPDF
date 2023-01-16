using iText.Signatures;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace MergeStudentPDF2.Model
{
    public class StudentInfo
    {
        public int SID { get; set; }
   
        public string DepartmentCode { get; set; }
        
        public string Reg_Num { get; set; }


        public StudentInfo()
        {
            SID = -1;

            DepartmentCode = string.Empty;

            Reg_Num = string.Empty;
        } // end StudentInfo
    } // end StudentInfo
}
