using Dapper;
using iText.Kernel.Geom;
using MergeStudentPDF2.Model;
using MySql.Data.MySqlClient;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace MergeStudentPDF2.Librarys
{
    public class EPItemTask
    {
        private static string _DataParentDirect = string.Empty;


        protected EPItemTask()
        { 
        }

        public static void SetDataParentDirect(string Direct)
        {
            _DataParentDirect = Direct;
        } // end SetDataParentDirect


        public static List<StudentInfo> GetStudentSIDList()
        {
            try
            {
                List<StudentInfo> SIDList;

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    string CommandText = "SELECT students.sid, students.reg_num, class.name as departmentcode FROM students, co, class WHERE co.oid = students.sid AND class.cid = co.cid ORDER BY sid";

                    SIDList = Conn.Query<StudentInfo>(CommandText).AsList();
                } // end using

                return SIDList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentSIDList


        public static JObject GetStudentEPFileList(int SID)
        {
            try
            {
                JObject EPFileList = new JObject()
                {
                    { "FormatType", 0 },

                    { "InfoEFID", 0 },

                    { "InfoFileType", 0 },

                    { "CourseEFID", 0 },

                    { "CourseFileType", 0 },

                    { "CourseLearnResultEFID", 0 },

                    { "CourseLearnResultFileType", 0 },

                    { "MultiplePerformanceEFID", 0 },

                    { "MultiplePerformanceFileType", 0 },

                    { "ComprehensiveExperienceEFID", 0 },

                    { "ComprehensiveExperienceFileType", 0 },

                    { "ProcessEFID", 0 },

                    { "ProcessFileType", 0 },

                    { "OtherEFID", 0 },

                    { "OtherFileType", 0 },

                    { "Semester6EFID", 0 },

                    { "Semester6FileType", 0 },

                    { "ProfessionalEFID", 0 },

                    { "ProfessionalFileType", 0 },

                    { "MergePDFEFID", 0 },

                    { "MergePDFileType", 0 },
                };

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_other_documents WHERE sid = @SID";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int FormatType = Reader.GetValue(Reader.GetOrdinal("format_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("format_type"));

                            int InfoEFID = Reader.GetValue(Reader.GetOrdinal("info_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("info_EFID"));

                            int InfoFileType = Reader.GetValue(Reader.GetOrdinal("info_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("info_file_type"));

                            int CourseEFID = Reader.GetValue(Reader.GetOrdinal("course_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("course_EFID"));

                            int CourseFileType = Reader.GetValue(Reader.GetOrdinal("course_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("course_file_type"));

                            int CourseLearnResultEFID = Reader.GetValue(Reader.GetOrdinal("learning_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("learning_EFID"));

                            int CourseLearnResultFileType = Reader.GetValue(Reader.GetOrdinal("learning_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("learning_file_type"));

                            int MultiplePerformanceEFID = Reader.GetValue(Reader.GetOrdinal("performance_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("performance_EFID"));

                            int MultiplePerformanceFileType = Reader.GetValue(Reader.GetOrdinal("performance_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("performance_file_type"));

                            int ComprehensiveExperienceEFID = Reader.GetValue(Reader.GetOrdinal("multiple_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("multiple_EFID"));

                            int ComprehensiveExperienceFileType = Reader.GetValue(Reader.GetOrdinal("multiple_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("multiple_file_type"));

                            int ProcessEFID = Reader.GetValue(Reader.GetOrdinal("process_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("process_EFID"));

                            int ProcessFileType = Reader.GetValue(Reader.GetOrdinal("process_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("process_file_type"));

                            int OtherEFID = Reader.GetValue(Reader.GetOrdinal("others_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("others_EFID"));

                            int OtherFileType = Reader.GetValue(Reader.GetOrdinal("others_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("others_file_type"));

                            int Semester6EFID = Reader.GetValue(Reader.GetOrdinal("semester6_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("semester6_EFID"));

                            int Semester6FileType = Reader.GetValue(Reader.GetOrdinal("semester6_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("semester6_file_type"));

                            int ProfessionalEFID = Reader.GetValue(Reader.GetOrdinal("implement_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("implement_EFID"));

                            int ProfessionalFileType = Reader.GetValue(Reader.GetOrdinal("implement_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("implement_file_type"));

                            int MergePDFEFID = Reader.GetValue(Reader.GetOrdinal("mergepdf_EFID")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("mergepdf_EFID"));

                            int MergePDFileType = Reader.GetValue(Reader.GetOrdinal("mergepdf_file_type")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("mergepdf_file_type"));

                            EPFileList["FormatType"] = FormatType;

                            EPFileList["InfoEFID"] = InfoEFID;

                            EPFileList["InfoFileType"] = InfoFileType;

                            EPFileList["CourseEFID"] = CourseEFID;

                            EPFileList["CourseFileType"] = CourseFileType;

                            EPFileList["CourseLearnResultEFID"] = CourseLearnResultEFID;

                            EPFileList["CourseLearnResultFileType"] = CourseLearnResultFileType;

                            EPFileList["MultiplePerformanceEFID"] = MultiplePerformanceEFID;

                            EPFileList["MultiplePerformanceFileType"] = MultiplePerformanceFileType;

                            EPFileList["ComprehensiveExperienceEFID"] = ComprehensiveExperienceEFID;

                            EPFileList["ComprehensiveExperienceFileType"] = ComprehensiveExperienceFileType;

                            EPFileList["ProcessEFID"] = ProcessEFID;

                            EPFileList["ProcessFileType"] = ProcessFileType;

                            EPFileList["OtherEFID"] = OtherEFID;

                            EPFileList["OtherFileType"] = OtherFileType;

                            EPFileList["Semester6EFID"] = Semester6EFID;

                            EPFileList["Semester6FileType"] = Semester6FileType;

                            EPFileList["ProfessionalEFID"] = ProfessionalEFID;

                            EPFileList["ProfessionalFileType"] = ProfessionalFileType;

                            EPFileList["MergePDFEFID"] = MergePDFEFID;

                            EPFileList["MergePDFileType"] = MergePDFileType;
                        } // end while
                    } // end if
                } // end using

                return EPFileList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentEPFileList


        public static JObject GetStudentBaseInfo(int SID)
        {
            try
            {
                JObject BaseInfo = new JObject();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = @"SELECT students.reg_num,
                                                   obj.name,
                                                   (SELECT IFNULL(students_info.school_name, highschool.name)) as schoolname,
                                                   (SELECT IFNULL(county.name, highschool.county)) as county,
                                                   students.cand_status,
                                                   students_info.institute_name,
                                                   students_info.class_name,
                                                   students_info.group_name,
                                                   students_info.subject_name,
                                                   students_info.experimental_type,
                                                   students.low_income,
                                                   students.cand_status,
                                                   class.name as classname
                                            FROM students
                                            LEFT JOIN students_info ON students_info.sid = students.sid
                                            LEFT JOIN obj ON obj.oid = students.sid
                                            LEFT JOIN highschool ON highschool.hsid = students.hsid
                                            LEFT JOIN county ON students_info.school_county_code = county.code
                                            LEFT JOIN co ON students.sid = co.oid
                                            LEFT JOIN class ON co.cid = class.cid
                                            WHERE students.sid = @SID";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string RegNum = Reader.GetValue(Reader.GetOrdinal("reg_num")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("reg_num")).ToString();

                            string Name = Reader.GetValue(Reader.GetOrdinal("name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("name")).ToString();

                            string SchoolName = Reader.GetValue(Reader.GetOrdinal("schoolname")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("schoolname")).ToString();

                            string County = Reader.GetValue(Reader.GetOrdinal("county")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("county")).ToString();

                            string InstituteName = Reader.GetValue(Reader.GetOrdinal("institute_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("institute_name")).ToString();

                            string ClassName = Reader.GetValue(Reader.GetOrdinal("class_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("class_name")).ToString();

                            string GroupName = Reader.GetValue(Reader.GetOrdinal("group_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("group_name")).ToString();

                            string SubjectName = Reader.GetValue(Reader.GetOrdinal("subject_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("subject_name")).ToString();

                            string ExperimentalType = Reader.GetValue(Reader.GetOrdinal("experimental_type")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("experimental_type")).ToString();

                            string LowIncome = Reader.GetValue(Reader.GetOrdinal("low_income")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("low_income")).ToString();

                            string CandStatus = Reader.GetValue(Reader.GetOrdinal("cand_status")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("cand_status")).ToString();

                            string DepartmentCode = Reader.GetValue(Reader.GetOrdinal("classname")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("classname")).ToString();

                            BaseInfo.Add("RegNum", RegNum);

                            BaseInfo.Add("Name", Name);

                            BaseInfo.Add("SchoolName", SchoolName);

                            BaseInfo.Add("County", County);

                            BaseInfo.Add("InstituteName", InstituteName);

                            BaseInfo.Add("ClassName", ClassName);

                            BaseInfo.Add("GroupName", GroupName);

                            BaseInfo.Add("SubjectName", SubjectName);

                            BaseInfo.Add("ExperimentalType", ExperimentalType);

                            BaseInfo.Add("DepartmentCode", DepartmentCode);

                            BaseInfo.Add("LowIncome", LowIncome);

                            BaseInfo.Add("CandStatus", CandStatus);
                        } // end while
                    } // end if
                } // end using

                return BaseInfo;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentBaseInfo

        
        public static JArray GetStudentFiveSemesterGrade(int SID)
        {
            try
            {
                JArray SemesterGradeList = new JArray();

                JArray TempSemesterGradeList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = @"SELECT * FROM students_course_record_semestergrade WHERE sid = @SID ORDER BY school_year, semester";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int SchoolYear = Reader.GetValue(Reader.GetOrdinal("school_year")) is DBNull ? -999 : Reader.GetInt32(Reader.GetOrdinal("school_year"));

                            int Semester = Reader.GetValue(Reader.GetOrdinal("semester")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("semester"));

                            decimal SemesterGrade = Reader.GetValue(Reader.GetOrdinal("semester_grade")) is DBNull ? -1 : Reader.GetDecimal(Reader.GetOrdinal("semester_grade"));

                            JObject SchoolGradeInfo = Reader.GetValue(Reader.GetOrdinal("school_grade_info")) is DBNull ? null : JObject.Parse(Reader.GetValue(Reader.GetOrdinal("school_grade_info")).ToString());

                            JObject GroupGradeInfo = Reader.GetValue(Reader.GetOrdinal("group_grade_info")) is DBNull ? null : JObject.Parse(Reader.GetValue(Reader.GetOrdinal("group_grade_info")).ToString());

                            JObject SubjectGradeInfo = Reader.GetValue(Reader.GetOrdinal("subject_grade_info")) is DBNull ? null : JObject.Parse(Reader.GetValue(Reader.GetOrdinal("subject_grade_info")).ToString());

                            JObject SemesterGradeObj = new JObject();

                            SemesterGradeObj.Add("SchoolYear", SchoolYear);

                            SemesterGradeObj.Add("Semester", Semester);

                            SemesterGradeObj.Add("SemesterGrade", SemesterGrade);

                            SemesterGradeObj.Add("SchoolGradeInfo", SchoolGradeInfo);

                            SemesterGradeObj.Add("GroupGradeInfo", GroupGradeInfo);

                            SemesterGradeObj.Add("SubjectGradeInfo", SubjectGradeInfo);

                            TempSemesterGradeList.Add(SemesterGradeObj);
                        } // end while
                    } // end if
                } // end using

                for (int i = 0; i < TempSemesterGradeList.Count; i++)
                {
                    JObject SemesterGradeObj = JObject.Parse(TempSemesterGradeList[i].ToString());

                    decimal SchoolRelativePerformance = GetRelativePerformance(JObject.Parse(SemesterGradeObj["SchoolGradeInfo"].ToString()), Convert.ToDecimal(SemesterGradeObj["SemesterGrade"].ToString()));

                    decimal GroupRelativePerformance = GetRelativePerformance(JObject.Parse(SemesterGradeObj["GroupGradeInfo"].ToString()), Convert.ToDecimal(SemesterGradeObj["SemesterGrade"].ToString()));

                    decimal SubjectRelativePerformance = GetRelativePerformance(JObject.Parse(SemesterGradeObj["SubjectGradeInfo"].ToString()), Convert.ToDecimal(SemesterGradeObj["SemesterGrade"].ToString()));

                    SemesterGradeObj.Add("SchoolRelativePerformance", SchoolRelativePerformance);

                    SemesterGradeObj.Add("GroupRelativePerformance", GroupRelativePerformance);

                    SemesterGradeObj.Add("SubjectRelativePerformance", SubjectRelativePerformance);

                    SemesterGradeList.Add(SemesterGradeObj);
                } // end for

                return SemesterGradeList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentFiveSemesterGrade


        public static JObject GetStudentCourseRecordList(int SID)
        {
            try
            {
                Dictionary<string, JArray> CourseRecordDictionary = new Dictionary<string, JArray>();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_course_record_learnrecord WHERE sid = @SID ORDER BY school_year, semester";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int SchoolYear = Reader.GetValue(Reader.GetOrdinal("school_year")) is DBNull ? -999 : Reader.GetInt32(Reader.GetOrdinal("school_year"));

                            int Semester = Reader.GetValue(Reader.GetOrdinal("semester")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("semester"));

                            string CourseName = Reader.GetValue(Reader.GetOrdinal("course_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("course_name")).ToString();

                            string CourseCategory = Reader.GetValue(Reader.GetOrdinal("course_category")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("course_category")).ToString();

                            string CourseCredits = Reader.GetValue(Reader.GetOrdinal("course_credits")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("course_credits")).ToString();

                            string Score = Reader.GetValue(Reader.GetOrdinal("score")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("score")).ToString();

                            string Interval = Reader.GetValue(Reader.GetOrdinal("course_score_interval")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("course_score_interval")).ToString();

                            int MotherPopulation = Reader.GetValue(Reader.GetOrdinal("course_population")) is DBNull ? 0 : Reader.GetInt32(Reader.GetOrdinal("course_population"));

                            JObject CourseObj = new JObject();

                            CourseObj.Add("SchoolYear", SchoolYear);

                            CourseObj.Add("Semester", Semester);

                            CourseObj.Add("CourseName", CourseName);

                            CourseObj.Add("CourseCredits", CourseCredits);

                            CourseObj.Add("CourseCategory", CourseCategory);

                            CourseObj.Add("Score", Score);

                            if (Interval != string.Empty)
                            {
                                JObject IntervalObj = JObject.Parse(Interval);

                                IntervalObj = new JObject(IntervalObj.Properties().OrderBy(x => Convert.ToInt32(x.Name.Split("_")[0].Substring(2))));

                                JArray IntervalList = GetIntervalValueList(IntervalObj);

                                decimal RelativePerformance = CalculateReletivePerformance(decimal.Parse(Score), IntervalList, MotherPopulation);

                                CourseObj.Add("RelativePerformance", RelativePerformance);
                            }
                            else 
                            {
                                CourseObj.Add("RelativePerformance", -1);
                            } // end if

                            string SchoolYearSemester = string.Format("{0}-{1}", SchoolYear, Semester);

                            if (!CourseRecordDictionary.ContainsKey(SchoolYearSemester))
                            {
                                CourseRecordDictionary[SchoolYearSemester] = new JArray();
                            } // end if

                            CourseRecordDictionary[SchoolYearSemester].Add(CourseObj);
                        } // end while
                    } // end if
                } // end using

                return JObject.FromObject(CourseRecordDictionary);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentCourseList


        public static JArray GetStudentCourseLearnResult(int SID)
        {
            try
            {
                JArray CourseLearnResultList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_courselearnresult, experience_files WHERE students_courselearnresult.sid = @SID AND experience_files.efid = students_courselearnresult.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int SchoolYear = Reader.GetValue(Reader.GetOrdinal("school_year")) is DBNull ? -999 : Reader.GetInt32(Reader.GetOrdinal("school_year"));

                            int Semester = Reader.GetValue(Reader.GetOrdinal("semester")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("semester"));

                            string CourseName = Reader.GetValue(Reader.GetOrdinal("course_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("course_name")).ToString();

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            int AchievementType = Reader.GetValue(Reader.GetOrdinal("achievement_type")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("achievement_type"));

                            JObject CourseLearnResult = new JObject();

                            CourseLearnResult.Add("SchoolYear", SchoolYear);

                            CourseLearnResult.Add("Semester", Semester);

                            CourseLearnResult.Add("CourseName", CourseName);

                            CourseLearnResult.Add("Desc", Desc);

                            CourseLearnResult.Add("AchievementType", AchievementType);

                            CourseLearnResult.Add("DocUrl", DocUrl);

                            CourseLearnResult.Add("VideoUrl", VideoUrl);

                            CourseLearnResultList.Add(CourseLearnResult);
                        } // end while
                    } // end if
                } // end using

                return CourseLearnResultList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentCourseLearnResult


        public static JObject GetStudentMultiplePerformance(int SID, out int MulitiplePerformanceCount)
        {
            try
            {
                JObject MultiplePerformance = new JObject();

                int Count = 0;

                // 校傳幹部
                JArray StudentSchoolCadreList = GetStudentSchoolCadre(SID);

                Count += StudentSchoolCadreList.Count;

                MultiplePerformance.Add("SchoolCadre", StudentSchoolCadreList);

                // 幹部經歷既事蹟紀錄
                JArray StudentCadreList = GetStudentCadre(SID);

                Count += StudentCadreList.Count;

                MultiplePerformance.Add("Cadre", StudentCadreList);

                // 競賽
                JArray CompetitionList = GetStudentCompetition(SID);

                Count += CompetitionList.Count;

                MultiplePerformance.Add("Competition", CompetitionList);

                // 檢定證照
                JArray CertificationList = GetStudentCertification(SID);

                Count += CertificationList.Count;

                MultiplePerformance.Add("Certification", CertificationList);

                // 服務學習
                JArray ServiceLearnList = GetStudentServiceLearn(SID);

                Count += ServiceLearnList.Count;

                MultiplePerformance.Add("ServiceLearn", ServiceLearnList);

                // 彈性時間學習
                JArray AlternativeLearnList = GetStudentAlternativeLearn(SID);

                Count += AlternativeLearnList.Count;

                MultiplePerformance.Add("AlternativeLearn", AlternativeLearnList);

                // 職場學習
                JArray JobLearnList = GetStudentJobLearn(SID);

                Count += JobLearnList.Count;

                MultiplePerformance.Add("JobLearn", JobLearnList);

                // 作品成果
                JArray CollectionList = GetStudentCollection(SID);

                Count += CollectionList.Count;

                MultiplePerformance.Add("Collection", CollectionList);

                // 團體活動
                JArray TeamActivityList = GetStudentTeamActivity(SID);

                Count += TeamActivityList.Count;

                MultiplePerformance.Add("TeamActivity", TeamActivityList);

                // 其他多元表現
                JArray OtherMultiplePerformanceList = GetStudentOtherMultiplePerformance(SID);

                Count += OtherMultiplePerformanceList.Count;

                MultiplePerformance.Add("OtherMultiplePerformance", OtherMultiplePerformanceList);

                // 大學先修
                JArray AdvancedPlacementList = GetStudentAdvancedPlacement(SID);

                Count += AdvancedPlacementList.Count;

                MultiplePerformance.Add("AdvancedPlacement", AdvancedPlacementList);

                MulitiplePerformanceCount = Count;

                return MultiplePerformance;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentMultiplePerformance


        public static JArray GetStudentSchoolCadre(int SID)
        {
            try
            {
                JArray SchoolCadreList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_schoolcadre WHERE sid = @SID";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            string PositionName = Reader.GetValue(Reader.GetOrdinal("position_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("position_name")).ToString();

                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            JObject SchoolCadreObj = new JObject();

                            SchoolCadreObj.Add("MajorName", MajorName);

                            SchoolCadreObj.Add("StartDate", StartDate);

                            SchoolCadreObj.Add("EndDate", EndDate);

                            SchoolCadreObj.Add("PositionName", PositionName);

                            SchoolCadreObj.Add("LevelID", LevelID);

                            SchoolCadreList.Add(SchoolCadreObj);
                        } // end while
                    } // end if
                } // end using

                return SchoolCadreList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try cathc
        } // end GetStudentSchoolCadre


        public static JArray GetStudentCadre(int SID)
        {
            try
            {
                JArray CadreList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_cadre, experience_files WHERE students_multipleperformance_cadre.sid = @SID AND experience_files.efid = students_multipleperformance_cadre.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            string PositionName = Reader.GetValue(Reader.GetOrdinal("position_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("position_name")).ToString();

                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject CadreObj = new JObject();

                            CadreObj.Add("MajorName", MajorName);

                            CadreObj.Add("StartDate", StartDate);

                            CadreObj.Add("EndDate", EndDate);

                            CadreObj.Add("PositionName", PositionName);

                            CadreObj.Add("LevelID", LevelID);

                            CadreObj.Add("Desc", Desc);

                            CadreObj.Add("DocUrl", DocUrl);

                            CadreObj.Add("VideoUrl", VideoUrl);

                            CadreObj.Add("VideoUrlExternal", VideoUrlExternal);

                            CadreList.Add(CadreObj);
                        } // end while
                    } // end if
                } // end using

                return CadreList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentCadre


        public static JArray GetStudentCompetition(int SID)
        {
            try
            {
                JArray CompetitionList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_competition, experience_files WHERE students_multipleperformance_competition.sid = @SID AND experience_files.efid = students_multipleperformance_competition.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            string Awards = Reader.GetValue(Reader.GetOrdinal("awards")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("awards")).ToString();

                            string Items = Reader.GetValue(Reader.GetOrdinal("items")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("items")).ToString();

                            string Announce_Date = Reader.GetValue(Reader.GetOrdinal("announce_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("announce_date")).ToString();

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject CompetitionObj = new JObject();

                            CompetitionObj.Add("MajorName", MajorName);

                            CompetitionObj.Add("LevelID", LevelID);

                            CompetitionObj.Add("Awards", Awards);

                            CompetitionObj.Add("Items", Items);

                            CompetitionObj.Add("Announce_Date", Announce_Date);

                            CompetitionObj.Add("Desc", Desc);

                            CompetitionObj.Add("DocUrl", DocUrl);

                            CompetitionObj.Add("VideoUrl", VideoUrl);

                            CompetitionObj.Add("VideoUrlExternal", VideoUrlExternal);

                            CompetitionList.Add(CompetitionObj);
                        } // end while
                    } // end if
                } // end using

                return CompetitionList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentCompetition


        public static JArray GetStudentCertification(int SID)
        {
            try
            {
                JArray CertificationList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_certification, experience_files WHERE students_multipleperformance_certification.sid = @SID AND experience_files.efid = students_multipleperformance_certification.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string LicenseInfo = Reader.GetValue(Reader.GetOrdinal("license_info")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("license_info")).ToString();

                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            decimal Score = Reader.GetValue(Reader.GetOrdinal("score")) is DBNull ? -999.99M : Reader.GetDecimal(Reader.GetOrdinal("score"));

                            string ScoreList = Reader.GetValue(Reader.GetOrdinal("score_list")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("score_list")).ToString();

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject CertificationObj = new JObject();

                            CertificationObj.Add("LicenseInfo", LicenseInfo == string.Empty ? null : JObject.Parse(LicenseInfo));

                            CertificationObj.Add("LevelID", LevelID);

                            CertificationObj.Add("Score", Score);

                            CertificationObj.Add("ScoreList", ScoreList);

                            CertificationObj.Add("Desc", Desc);

                            CertificationObj.Add("DocUrl", DocUrl);

                            CertificationObj.Add("VideoUrl", VideoUrl);

                            CertificationObj.Add("VideoUrlExternal", VideoUrlExternal);

                            CertificationList.Add(CertificationObj);
                        } // end while
                    } // end if
                } // end using

                return CertificationList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentCertification


        public static JArray GetStudentServiceLearn(int SID)
        {
            try
            {
                JArray ServiceLearnList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_servicelearn, experience_files WHERE students_multipleperformance_servicelearn.sid = @SID AND experience_files.efid = students_multipleperformance_servicelearn.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string ServiceUnit = Reader.GetValue(Reader.GetOrdinal("service_unit")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("service_unit")).ToString();

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            decimal ServiceHours = Reader.GetValue(Reader.GetOrdinal("service_hours")) is DBNull ? -9999.999M : Reader.GetDecimal(Reader.GetOrdinal("service_hours"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject ServiceLearnObj = new JObject();

                            ServiceLearnObj.Add("MajorName", MajorName);

                            ServiceLearnObj.Add("ServiceUnit", ServiceUnit);

                            ServiceLearnObj.Add("StartDate", StartDate);

                            ServiceLearnObj.Add("EndDate", EndDate);

                            ServiceLearnObj.Add("ServiceHours", ServiceHours);

                            ServiceLearnObj.Add("Desc", Desc);

                            ServiceLearnObj.Add("DocUrl", DocUrl);

                            ServiceLearnObj.Add("VideoUrl", VideoUrl);

                            ServiceLearnObj.Add("VideoUrlExternal", VideoUrlExternal);

                            ServiceLearnList.Add(ServiceLearnObj);
                        } // end while
                    } // end if
                } // end using

                return ServiceLearnList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentServiceLearn


        public static JArray GetStudentAlternativeLearn(int SID)
        {
            try
            {
                JArray AlternativeLearnList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_alternativelearn, experience_files WHERE students_multipleperformance_alternativelearn.sid = @SID AND experience_files.efid = students_multipleperformance_alternativelearn.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int SchoolYear = Reader.GetValue(Reader.GetOrdinal("school_year")) is DBNull ? -999 : Reader.GetInt32(Reader.GetOrdinal("school_year"));

                            int Semester = Reader.GetValue(Reader.GetOrdinal("semester")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("semester"));

                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string AlternativeUnit = Reader.GetValue(Reader.GetOrdinal("alternative_unit")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("alternative_unit")).ToString();

                            int SessionPerWeek = Reader.GetValue(Reader.GetOrdinal("session_per_week")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("session_per_week"));

                            int SessionWeeks = Reader.GetValue(Reader.GetOrdinal("session_weeks")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("session_weeks"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject AlternativeLearnObj = new JObject();

                            AlternativeLearnObj.Add("SchoolYear", SchoolYear);

                            AlternativeLearnObj.Add("Semester", Semester);

                            AlternativeLearnObj.Add("LevelID", LevelID);

                            AlternativeLearnObj.Add("MajorName", MajorName);

                            AlternativeLearnObj.Add("AlternativeUnit", AlternativeUnit);

                            AlternativeLearnObj.Add("SessionPerWeek", SessionPerWeek);

                            AlternativeLearnObj.Add("SessionWeeks", SessionWeeks);

                            AlternativeLearnObj.Add("Desc", Desc);

                            AlternativeLearnObj.Add("DocUrl", DocUrl);

                            AlternativeLearnObj.Add("VideoUrl", VideoUrl);

                            AlternativeLearnObj.Add("VideoUrlExternal", VideoUrlExternal);

                            AlternativeLearnList.Add(AlternativeLearnObj);
                        } // end while
                    } // end if
                } // end using

                return AlternativeLearnList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentAlternativelearn


        public static JArray GetStudentJobLearn(int SID)
        {
            try
            {
                JArray JobLearnList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_joblearn, experience_files WHERE students_multipleperformance_joblearn.sid = @SID AND experience_files.efid = students_multipleperformance_joblearn.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            string JobUnit = Reader.GetValue(Reader.GetOrdinal("job_unit")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("job_unit")).ToString();

                            string JobTitle = Reader.GetValue(Reader.GetOrdinal("job_title")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("job_title")).ToString();

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            decimal JobHours = Reader.GetValue(Reader.GetOrdinal("job_hours")) is DBNull ? -9999.999M : Reader.GetInt32(Reader.GetOrdinal("job_hours"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject JobLearnObj = new JObject();

                            JobLearnObj.Add("LevelID", LevelID);

                            JobLearnObj.Add("JobUnit", JobUnit);

                            JobLearnObj.Add("JobTitle", JobTitle);

                            JobLearnObj.Add("StartDate", StartDate);

                            JobLearnObj.Add("EndDate", EndDate);

                            JobLearnObj.Add("JobHours", JobHours);

                            JobLearnObj.Add("Desc", Desc);

                            JobLearnObj.Add("DocUrl", DocUrl);

                            JobLearnObj.Add("VideoUrl", VideoUrl);

                            JobLearnObj.Add("VideoUrlExternal", VideoUrlExternal);

                            JobLearnList.Add(JobLearnObj);
                        } // end while
                    } // end if
                } // end using

                return JobLearnList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentJobLearn


        public static JArray GetStudentCollection(int SID)
        {
            try
            {
                JArray CollectionList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_collection, experience_files WHERE students_multipleperformance_collection.sid = @SID AND experience_files.efid = students_multipleperformance_collection.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string CollectionDate = Reader.GetValue(Reader.GetOrdinal("collection_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("collection_date")).ToString();

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject CollectionObj = new JObject();

                            CollectionObj.Add("MajorName", MajorName);

                            CollectionObj.Add("CollectionDate", CollectionDate);

                            CollectionObj.Add("Desc", Desc);

                            CollectionObj.Add("DocUrl", DocUrl);

                            CollectionObj.Add("VideoUrl", VideoUrl);

                            CollectionObj.Add("VideoUrlExternal", VideoUrlExternal);

                            CollectionList.Add(CollectionObj);
                        } // end while
                    } // end if
                } // end using

                return CollectionList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentCollection


        public static JArray GetStudentTeamActivity(int SID)
        {
            try
            {
                JArray TeamActivityList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_teamactivity, experience_files WHERE students_multipleperformance_teamactivity.sid = @SID AND experience_files.efid = students_multipleperformance_teamactivity.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            int SchoolYear = Reader.GetValue(Reader.GetOrdinal("school_year")) is DBNull ? -999 : Reader.GetInt32(Reader.GetOrdinal("school_year"));

                            int Semester = Reader.GetValue(Reader.GetOrdinal("semester")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("semester"));

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            int LevelID = Reader.GetValue(Reader.GetOrdinal("level_id")) is DBNull ? -1 : Reader.GetInt32(Reader.GetOrdinal("level_id"));

                            string ActivityUnit = Reader.GetValue(Reader.GetOrdinal("activity_unit")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("activity_unit")).ToString();

                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            decimal ActivityHours = Reader.GetValue(Reader.GetOrdinal("activity_hours")) is DBNull ? -1 : Reader.GetDecimal(Reader.GetOrdinal("activity_hours"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject TeamActivityObj = new JObject();

                            TeamActivityObj.Add("SchoolYear", SchoolYear);

                            TeamActivityObj.Add("Semester", Semester);

                            TeamActivityObj.Add("StartDate", StartDate);

                            TeamActivityObj.Add("EndDate", EndDate);

                            TeamActivityObj.Add("LevelID", LevelID);

                            TeamActivityObj.Add("ActivityUnit", ActivityUnit);

                            TeamActivityObj.Add("MajorName", MajorName);

                            TeamActivityObj.Add("ActivityHours", ActivityHours);

                            TeamActivityObj.Add("Desc", Desc);

                            TeamActivityObj.Add("DocUrl", DocUrl);

                            TeamActivityObj.Add("VideoUrl", VideoUrl);

                            TeamActivityObj.Add("VideoUrlExternal", VideoUrlExternal);

                            TeamActivityList.Add(TeamActivityObj);
                        } // end while
                    } // end if
                } // end using

                return TeamActivityList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentTeamActivity


        public static JArray GetStudentOtherMultiplePerformance(int SID)
        {
            try
            {
                JArray OtherMultiplePerformanceList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_others, experience_files WHERE students_multipleperformance_others.sid = @SID AND experience_files.efid = students_multipleperformance_others.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string MultipleUnit = Reader.GetValue(Reader.GetOrdinal("multiple_unit")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("multiple_unit")).ToString();

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            decimal MultipleHours = Reader.GetValue(Reader.GetOrdinal("multiple_hours")) is DBNull ? -9999.999M : Reader.GetDecimal(Reader.GetOrdinal("multiple_hours"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject OtherMultiplePerformanceObj = new JObject();

                            OtherMultiplePerformanceObj.Add("MajorName", MajorName);

                            OtherMultiplePerformanceObj.Add("MultipleUnit", MultipleUnit);

                            OtherMultiplePerformanceObj.Add("StartDate", StartDate);

                            OtherMultiplePerformanceObj.Add("EndDate", EndDate);

                            OtherMultiplePerformanceObj.Add("MultipleHours", MultipleHours);

                            OtherMultiplePerformanceObj.Add("Desc", Desc);

                            OtherMultiplePerformanceObj.Add("DocUrl", DocUrl);

                            OtherMultiplePerformanceObj.Add("VideoUrl", VideoUrl);

                            OtherMultiplePerformanceObj.Add("VideoUrlExternal", VideoUrlExternal);

                            OtherMultiplePerformanceList.Add(OtherMultiplePerformanceObj);
                        } // end while
                    } // end if
                } // end using

                return OtherMultiplePerformanceList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentOtherMultiplePerformance


        public static JArray GetStudentAdvancedPlacement(int SID)
        {
            try
            {
                JArray AdvancedPlacementList = new JArray();

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM students_multipleperformance_advancedplacement, experience_files WHERE students_multipleperformance_advancedplacement.sid = @SID AND experience_files.efid = students_multipleperformance_advancedplacement.efid";

                    Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            string MajorName = Reader.GetValue(Reader.GetOrdinal("major_name")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("major_name")).ToString();

                            string AdvancedUnit = Reader.GetValue(Reader.GetOrdinal("advanced_unit")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("advanced_unit")).ToString();

                            string AdvancedCourse = Reader.GetValue(Reader.GetOrdinal("advanced_course")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("advanced_course")).ToString();

                            string StartDate = Reader.GetValue(Reader.GetOrdinal("start_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("start_date")).ToString();

                            string EndDate = Reader.GetValue(Reader.GetOrdinal("end_date")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("end_date")).ToString();

                            decimal AdvancedCredits = Reader.GetValue(Reader.GetOrdinal("advanced_credits")) is DBNull ? 0 : Reader.GetDecimal(Reader.GetOrdinal("advanced_credits"));

                            decimal AdvancedHours = Reader.GetValue(Reader.GetOrdinal("advanced_hours")) is DBNull ? -9999.999M : Reader.GetDecimal(Reader.GetOrdinal("advanced_hours"));

                            string Desc = Reader.GetValue(Reader.GetOrdinal("desc")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("desc")).ToString();

                            string DocUrl = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();

                            string VideoUrl = Reader.GetValue(Reader.GetOrdinal("video_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url")).ToString();

                            string VideoUrlExternal = Reader.GetValue(Reader.GetOrdinal("video_url_external")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("video_url_external")).ToString();

                            JObject AdvancedPlacementObj = new JObject();

                            AdvancedPlacementObj.Add("MajorName", MajorName);

                            AdvancedPlacementObj.Add("AdvancedUnit", AdvancedUnit);

                            AdvancedPlacementObj.Add("AdvancedCourse", AdvancedCourse);

                            AdvancedPlacementObj.Add("StartDate", StartDate);

                            AdvancedPlacementObj.Add("EndDate", EndDate);

                            AdvancedPlacementObj.Add("AdvancedCredits", AdvancedCredits);

                            AdvancedPlacementObj.Add("AdvancedHours", AdvancedHours);

                            AdvancedPlacementObj.Add("Desc", Desc);

                            AdvancedPlacementObj.Add("DocUrl", DocUrl);

                            AdvancedPlacementObj.Add("VideoUrl", VideoUrl);

                            AdvancedPlacementObj.Add("VideoUrlExternal", VideoUrlExternal);

                            AdvancedPlacementList.Add(AdvancedPlacementObj);
                        } // end while
                    } // end if
                } // end using

                return AdvancedPlacementList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentAdvancedPlacement


        public static string GetStudentDocFilePathByEFID(int EFID)
        {
            try
            {
                string FileName = string.Empty;

                using (MySqlConnection Conn = DBHelper.GetSqlConnection())
                {
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    } // end if

                    MySqlCommand Command = Conn.CreateCommand();

                    Command.CommandText = "SELECT * FROM experience_files WHERE efid = @EFID";

                    Command.Parameters.Add("@EFID", MySqlDbType.Int32).Value = EFID;

                    MySqlDataReader Reader = Command.ExecuteReader();

                    if (Reader.HasRows)
                    {
                        while (Reader.Read())
                        {
                            FileName = Reader.GetValue(Reader.GetOrdinal("doc_url")) is DBNull ? string.Empty : Reader.GetValue(Reader.GetOrdinal("doc_url")).ToString();
                        } // end while
                    } // end if
                } // end using

                return System.IO.Path.Combine(_DataParentDirect, FileName);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetStudentDocFilePathByEFID


        public static Func<JObject, string> GetMultiplePerformanceToHtmlFunc(string ItemName)
        {
            Func<JObject, string> FuncName = null;

            switch (ItemName)
            {
                case "Cadre":
                    FuncName = HtmlHelper.GenerateCadreHtmlString;
                    break;

                case "Competition":
                    FuncName = HtmlHelper.GenerateCompeititionHtmlString;
                    break;

                case "Certification":
                    FuncName = HtmlHelper.GenerateCertificationHtmlString;
                    break;

                case "ServiceLearn":
                    FuncName = HtmlHelper.GenerateServiceLearnHtmlString;
                    break;

                case "AlternativeLearn":
                    FuncName = HtmlHelper.GenerateAlternativeLearnHtmlString;
                    break;

                case "JobLearn":
                    FuncName = HtmlHelper.GenerateJobLearnHtmlString;
                    break;

                case "Collection":
                    FuncName = HtmlHelper.GenerateCollectionHtmlString;
                    break;

                case "TeamActivity":
                    FuncName = HtmlHelper.GenerateTeamActivityHtmlString;
                    break;

                case "OtherMultiplePerformance":
                    FuncName = HtmlHelper.GenerateOtherMultiplePerformanceHtmlString;
                    break;

                case "AdvancedPlacement":
                    FuncName = HtmlHelper.GenerateAdvancedPlacementHtmlString;
                    break;

                default:
                    FuncName = HtmlHelper.GenerateCadreHtmlString;
                    break;
            } // end switch

            return FuncName;
        } // end GetMultiplePerformanceToHtmlFunc


        public static string GetMultiplePerformanceItemName(string ItemName)
        {
            string Name = string.Empty;

            switch (ItemName)
            {
                case "SchoolCadre":
                    Name = "校⽅建⽴幹部經歷紀錄";
                    break;

                case "Cadre":
                    Name = "幹部經歷暨事蹟紀錄";
                    break;

                case "Competition":
                    Name = "競賽參與紀錄";
                    break;

                case "Certification":
                    Name = "檢定證照紀錄";
                    break;

                case "ServiceLearn":
                    Name = "服務學習紀錄";
                    break;

                case "AlternativeLearn":
                    Name = "彈性學習時間紀錄";
                    break;

                case "JobLearn":
                    Name = "職場學習紀錄";
                    break;

                case "Collection":
                    Name = "作品成果紀錄";
                    break;

                case "TeamActivity":
                    Name = "團體活動時間紀錄";
                    break;

                case "OtherMultiplePerformance":
                    Name = "其他多元表現紀錄";
                    break;

                case "AdvancedPlacement":
                    Name = "大學及技專校院先修課程紀錄";
                    break;

                default:
                    Name = "幹部經歷暨事蹟紀錄";
                    break;
            } // end switch

            return Name;
        } // end GetMultiplePerformanceToHtmlFunc


        public static string GetCadreLevelName(int Level)
        {
            string LevelName = string.Empty;

            switch (Level)
            {
                case 1:
                    LevelName = "校級幹部";
                    break;

                case 2:
                    LevelName = "班級幹部";
                    break;

                case 3:
                    LevelName = "社團幹部";
                    break;

                case 4:
                    LevelName = "實習幹部";
                    break;

                case 5:
                    LevelName = "校外自治組織團體";
                    break;

                case 9:
                    LevelName = "其他幹部";
                    break;

                default:
                    LevelName = "其他幹部";
                    break;
            } // end switch case

            return LevelName;
        } // end GetCadreLevelName


        public static string GetCompetitionLevelName(int Level)
        {
            string LevelName = string.Empty;

            switch (Level)
            {
                case 1:
                    LevelName = "校級";
                    break;

                case 2:
                    LevelName = "縣市級";
                    break;

                case 3:
                    LevelName = "全國";
                    break;

                case 4:
                    LevelName = "國際";
                    break;

                default:
                    LevelName = "校級";
                    break;
            } // end switch case

            return LevelName;
        } // end GetCompetitionLevelName


        public static string GetAlternativeLearnLevelName(int Level)
        {
            string LevelName = string.Empty;

            switch (Level)
            {
                case 1:
                    LevelName = "自主學習";
                    break;

                case 2:
                    LevelName = "選手培訓";
                    break;

                case 3:
                    LevelName = "充實(增廣)課程";
                    break;

                case 4:
                    LevelName = "補強性課程";
                    break;

                case 5:
                    LevelName = "學校特色活動";
                    break;

                default:
                    LevelName = "-";
                    break;
            } // end switch case

            return LevelName;
        } // end GetCompetitionLevelName


        public static string GetTeamActivityLevelName(int Level)
        {
            string LevelName = string.Empty;

            switch (Level)
            {
                case 1:
                    LevelName = "班級活動";
                    break;

                case 2:
                    LevelName = "社團活動";
                    break;

                case 3:
                    LevelName = "學生自治會活動";
                    break;

                case 4:
                    LevelName = "週會";
                    break;

                case 5:
                    LevelName = "講座";
                    break;

                case 6:
                    LevelName = "其他";
                    break;

                default:
                    LevelName = "-";
                    break;
            } // end switch case

            return LevelName;
        } // end GetCompetitionLevelName


        static decimal GetRelativePerformance(JObject GradeInfo, decimal SemesterGrade)
        {
            JObject IntervalObj = JObject.Parse(GradeInfo["分數級距分布"].ToString());

            IntervalObj = new JObject(IntervalObj.Properties().OrderBy(x => Convert.ToInt32(x.Name.Split("_")[0].Substring(2))));

            JArray IntervalValueList = GetIntervalValueList(IntervalObj);

            decimal RelativePerformance = CalculateReletivePerformance(SemesterGrade, IntervalValueList, Convert.ToInt32(GradeInfo["母體人數"].ToString()));

            return RelativePerformance;
        } // end GetRelativePerformance


        static JArray GetIntervalValueList(JObject IntervalObj)
        {
            JArray IntervalValueList = new JArray();

            foreach (JProperty jProperty in IntervalObj.Properties())
            {
                IntervalValueList.Add(Convert.ToInt32(IntervalObj[jProperty.Name].ToString()));
            } // end foreach

            return IntervalValueList;
        } // end GetIntervalValueList


        static decimal CalculateReletivePerformance(decimal Grade, JArray IntervalList, int MotherPopulation)
        {
            int Min = Grade == 100 ? 99 : (int)Math.Floor(Grade);

            int Max = Grade == 100 ? 100 : Math.Ceiling(Grade) == Math.Floor(Grade) ? Min + 1 : (int)Math.Ceiling(Grade);

            int RangeCount = Convert.ToInt32(IntervalList[Min].ToString());

            int LoseRangeCount = 0;

            for (int i = IntervalList.Count - 1; i > Min; i--)
            {
                LoseRangeCount += Convert.ToInt32(IntervalList[i].ToString());
            } // end for

            decimal Val1 = (1 - ((Grade - Min) / (Max - Min))) * RangeCount + LoseRangeCount;

            decimal ReletivePerformance = (Val1 / MotherPopulation) * 100;

            ReletivePerformance = Math.Round(ReletivePerformance, 1, MidpointRounding.AwayFromZero);

            return ReletivePerformance;
        } // end GetReletivePerformanceObj
    } // end EPItemTask
}
