using iText.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Navigation;
using iText.Kernel.Utils;
using iText.Layout;
using MySql.Data.MySqlClient;
using MySqlX.XDevAPI.Common;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeStudentPDF2.Librarys
{
    public class StudentPdfHandle
    {
        private static string _DataParnetDirect = string.Empty;


        public static void SetDataParnetDirect(string Direct)
        {
            _DataParnetDirect = Direct;

            EPItemTask.SetDataParentDirect(Direct);

            HtmlHelper.SetDataParentDirect(Direct);
        } // end SetDataParnetDirect


        public static void StartMerge(int SID)
        {
            try
            {
                JObject EPFileList = EPItemTask.GetStudentEPFileList(SID);

                int FormatType = Convert.ToInt32(EPFileList["FormatType"].ToString());

                JObject BaseInfo = EPItemTask.GetStudentBaseInfo(SID);

                int CourseLearnCount = 0;

                JArray CourseLearnResult = new JArray();

                int MultiplePerformanceCount = 0;

                JObject MultiplePerformance = new JObject();

                JArray FiveSemesterGradeInfoList = EPItemTask.GetStudentFiveSemesterGrade(SID);

                decimal FiveSemesterAvgGrade = FiveSemesterGradeInfoList.Count <= 0 ? -1 : Math.Round(FiveSemesterGradeInfoList.Average(x => Convert.ToDecimal(x["SemesterGrade"].ToString())), 1, MidpointRounding.AwayFromZero);

                JObject CourseRecordObj = EPItemTask.GetStudentCourseRecordList(SID);

                if (Convert.ToInt32(EPFileList["CourseLearnResultFileType"].ToString()) == 1)
                {
                    CourseLearnResult = EPItemTask.GetStudentCourseLearnResult(SID);

                    CourseLearnCount = CourseLearnResult.Count;
                }
                else if (Convert.ToInt32(EPFileList["CourseLearnResultFileType"].ToString()) == 2)
                {
                    CourseLearnCount = 1;

                    if (Convert.ToInt32(EPFileList["ProfessionalFileType"].ToString()) == 2)
                    {
                        CourseLearnCount += 1;
                    } // end if
                } // end if

                if (Convert.ToInt32(EPFileList["MultiplePerformanceFileType"].ToString()) == 1)
                {
                    MultiplePerformance = EPItemTask.GetStudentMultiplePerformance(SID, out MultiplePerformanceCount);
                }
                else if (Convert.ToInt32(EPFileList["MultiplePerformanceFileType"].ToString()) == 2)
                {
                    MultiplePerformanceCount = 1;
                } // end if

                List<string> FilePathList = new List<string>();

                if (FormatType == 1 || FormatType == 2)
                {
                    string DirectionPath = $"Merge/{BaseInfo["DepartmentCode"].ToString()}/{BaseInfo["RegNum"].ToString()}";

                    if (!Directory.Exists(DirectionPath))
                    {
                        Directory.CreateDirectory(DirectionPath);
                    } // end if

                    FilePathList.Add(GenerateBaseInfoPDFDocument(BaseInfo, FiveSemesterAvgGrade, CourseLearnCount, MultiplePerformanceCount, DirectionPath));

                    // 學習歷程自述
                    if (Convert.ToInt32(EPFileList["ProcessFileType"].ToString()) == 2 && FormatType > 0)
                    {
                        FilePathList.Add(GenerateProcessDoc(Convert.ToInt32(EPFileList["ProcessEFID"].ToString()), DirectionPath));
                    } // end if

                    // 修課紀錄
                    if (FormatType == 1)
                    {
                        FilePathList.Add(GenerateCourseRecordPdfDocument_RowData(FiveSemesterGradeInfoList, CourseRecordObj, DirectionPath));
                    }
                    else if (FormatType == 2)
                    {
                        if (Convert.ToInt32(EPFileList["CourseFileType"].ToString()) == 2)
                        {
                            FilePathList.Add(GenerateCourseRecordPdfDocument_PDF(Convert.ToInt32(EPFileList["CourseEFID"].ToString()), DirectionPath));
                        } // end if
                    } // end if

                    // 課程學習成果
                    if (Convert.ToInt32(EPFileList["CourseLearnResultFileType"].ToString()) == 1)
                    {
                        if (CourseLearnResult.Count > 0)
                        {
                            FilePathList.Add(GenerateCourseLearnResult_RowData(CourseLearnResult, DirectionPath));
                        } // end if
                    }
                    else if (Convert.ToInt32(EPFileList["CourseLearnResultFileType"].ToString()) == 2)
                    {
                        FilePathList.Add(GenerateCourseLearnResult_PDF(Convert.ToInt32(EPFileList["CourseLearnResultEFID"].ToString()), DirectionPath));

                        if (Convert.ToInt32(EPFileList["ProfessionalFileType"].ToString()) == 2)
                        {
                            FilePathList.Add(GenerateProfessionCourseLearnResult_PDF(Convert.ToInt32(EPFileList["ProfessionalEFID"].ToString()), DirectionPath));
                        } // end if
                    }
                    else 
                    {
                        if (Convert.ToInt32(EPFileList["ProfessionalFileType"].ToString()) == 2)
                        {
                            FilePathList.Add(GenerateProfessionCourseLearnResult_PDF(Convert.ToInt32(EPFileList["ProfessionalEFID"].ToString()), DirectionPath));
                        } // end if
                    } // end if

                    // 多元表現綜整心得
                    if (Convert.ToInt32(EPFileList["ComprehensiveExperienceFileType"].ToString()) == 2)
                    {
                        FilePathList.Add(GenerateComprehensiveExperienceDoc(Convert.ToInt32(EPFileList["ComprehensiveExperienceEFID"].ToString()), DirectionPath));
                    } // end if

                    // 多元表現
                    if (Convert.ToInt32(EPFileList["MultiplePerformanceFileType"].ToString()) == 1)
                    {
                        if (MultiplePerformanceCount > 0)
                        {
                            FilePathList.Add(GenerateMultiplePerformance(MultiplePerformance, DirectionPath));
                        } // end if
                    }
                    else if (Convert.ToInt32(EPFileList["MultiplePerformanceFileType"].ToString()) == 2)
                    {
                        FilePathList.Add(GenerateMultiplePerformance_PDF(Convert.ToInt32(EPFileList["MultiplePerformanceEFID"].ToString()), DirectionPath));
                    } // end if

                    // 其他
                    if (Convert.ToInt32(EPFileList["OtherFileType"].ToString()) == 2)
                    {
                        FilePathList.Add(GenerateOtherDoc_PDF(Convert.ToInt32(EPFileList["OtherEFID"].ToString()), DirectionPath));
                    } // end if

                    string OutputPath = Path.Combine(DirectionPath, $"{BaseInfo["RegNum"].ToString()}.pdf");

                    PdfHelper.MergeMultiple(FilePathList, OutputPath);

                    SetMergePDFPath(SID, OutputPath, Convert.ToInt32(EPFileList["MergePDFEFID"].ToString()));
                } // end if
            }
            catch (Exception ex)
            {
                SetMergeErrorEFID(SID);

                throw new Exception(ex.Message);
            } // end try catch
        } // end StartMerge


        // 基本資料
        public static string GenerateBaseInfoPDFDocument(JObject BaseInfo, decimal FiveSemesterAvgGrade, int CourseLearnCount, int MultiplePerformanceCount, string DirectionPath)
        {
            try
            {
                string BaseInfoHtml = HtmlHelper.GenerateBaseInfoHtmlString(BaseInfo, FiveSemesterAvgGrade, CourseLearnCount, MultiplePerformanceCount);

                string FilePath = Path.Combine(DirectionPath, "基本資料.pdf");

                string TempFilePath = Path.Combine(DirectionPath, "Temp基本資料.pdf");

                PdfHelper.RenderHtmlToPDF(BaseInfoHtml, TempFilePath);

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempFilePath), new PdfWriter(FilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("封面", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempFilePath))
                {
                    File.Delete(TempFilePath);
                } // end if

                return FilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateBaseInfoPDFDocument


        // 學習歷程自述
        public static string GenerateProcessDoc(int EFID, string DirectionPath)
        {
            try
            {
                string CoverFilePath = "合併PDF頁面/3-3 學習歷程自述.pdf";

                string ProcessFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string TempPDFFilePath = Path.Combine(DirectionPath, "學習歷程自述1.pdf");

                string PDFFilePath = Path.Combine(DirectionPath, "學習歷程自述.pdf");

                PdfHelper.MergePDF(CoverFilePath, ProcessFilePath, TempPDFFilePath);

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempPDFFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("學習歷程自述", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempPDFFilePath))
                {
                    File.Delete(TempPDFFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateProcessDoc


        // 修課紀錄
        public static string GenerateCourseRecordPdfDocument_RowData(JArray FiveSemesterGradeInfoList, JObject CourseRecordObj, string DirectionPath)
        {
            try
            {
                string CoverPDFFilePath = "合併PDF頁面/3-7 修課紀錄.pdf";

                string CourseRecordHtml = HtmlHelper.GetHeader();

                CourseRecordHtml += HtmlHelper.GenerateCourseRecordHtmlString(CourseRecordObj);

                if (FiveSemesterGradeInfoList.Count > 0)
                {
                    CourseRecordHtml += HtmlHelper.GenerateSemesterAverageGradeAndReletivePerformanceHtmlString(FiveSemesterGradeInfoList);
                } // end if

                CourseRecordHtml += HtmlHelper.GetFooter();

                string TempFilePath = Path.Combine(DirectionPath, "Temp修課紀錄.pdf");

                PdfHelper.RenderHtmlToPDF(CourseRecordHtml, TempFilePath);

                string TempPDFFilePath = Path.Combine(DirectionPath, "修課紀錄1.pdf");

                PdfHelper.MergePDF(CoverPDFFilePath, TempFilePath, TempPDFFilePath);

                if (File.Exists(TempFilePath))
                {
                    File.Delete(TempFilePath);
                } // end if

                string PDFFilePath = Path.Combine(DirectionPath, "修課紀錄.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempPDFFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("修課紀錄");

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetFirstPage()));

                    int Index = 2;

                    PdfOutline CourseListOutline = PDFOutline.AddOutline("修課清單及成績");

                    CourseListOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(Index)));

                    foreach (JProperty property in CourseRecordObj.Properties())
                    {
                        PdfOutline CourseListSubOutline = CourseListOutline.AddOutline(property.Name, -1);

                        CourseListSubOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(Index)));

                        CourseListOutline.GetAllChildren().Add(CourseListSubOutline);

                        Index++;
                    } // end foreach

                    PDFOutline.GetAllChildren().Add(CourseListOutline);

                    if (FiveSemesterGradeInfoList.Count > 0)
                    {
                        PdfOutline FiveSemesterCourseOutline = PDFOutline.AddOutline("學期平均成績及學習相對表現");

                        FiveSemesterCourseOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetLastPage()));

                        PDFOutline.GetAllChildren().Add(FiveSemesterCourseOutline);
                    } // end if
                } // end using

                if (File.Exists(TempPDFFilePath))
                {
                    File.Delete(TempPDFFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCourseRecordPdfDocument_RowData


        // 修課紀錄-PDF
        public static string GenerateCourseRecordPdfDocument_PDF(int EFID, string DirectionPath)
        {
            try
            {
                string CourseRecordFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string CoverFilePath = "合併PDF頁面/3-7 修課紀錄.pdf";

                string TempCourseRecordFilePath = Path.Combine(DirectionPath, "修課紀錄1.pdf");

                PdfHelper.MergePDF(CoverFilePath, CourseRecordFilePath, TempCourseRecordFilePath);

                string PDFFilePath = Path.Combine(DirectionPath, "修課紀錄.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempCourseRecordFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("修課紀錄", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempCourseRecordFilePath))
                {
                    File.Delete(TempCourseRecordFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GetCourseRecordPdfDocument


        // 課程學習成果
        public static string GenerateCourseLearnResult_RowData(JArray CourseLearnResultList, string DirectionPath)
        {
            try
            {
                List<string> CourseLearnResultFilePath = new List<string>();

                Dictionary<int, int> PageCountDic = new Dictionary<int, int>();

                for (int i = 0; i < CourseLearnResultList.Count; i++)
                {
                    string Html = HtmlHelper.GetHeader() + HtmlHelper.GenerateCourseLearnResultHtmlString(JObject.Parse(CourseLearnResultList[i].ToString())) + HtmlHelper.GetFooter();

                    string TempCourseLearnResultFilePath = Path.Combine(DirectionPath, $"CourseLearnResult{i}.pdf");

                    if (Path.GetExtension(CourseLearnResultList[i]["DocUrl"].ToString()).ToLower() == ".pdf")
                    {
                        string TempFilePath = Path.Combine(DirectionPath, $"CourseLearnResultTemp{i}.pdf");

                        PdfHelper.RenderHtmlToPDF(Html, TempFilePath);

                        PdfHelper.MergePDF(TempFilePath, Path.Combine(_DataParnetDirect, CourseLearnResultList[i]["DocUrl"].ToString()), TempCourseLearnResultFilePath, false, true);

                        if (File.Exists(TempFilePath))
                        {
                            File.Delete(TempFilePath);
                        } // end if
                    }
                    else
                    {
                        PdfHelper.RenderHtmlToPDF(Html, TempCourseLearnResultFilePath);
                    } // end if

                    CourseLearnResultFilePath.Add(TempCourseLearnResultFilePath);

                    using (PdfDocument PDF = new PdfDocument(new PdfReader(TempCourseLearnResultFilePath).SetUnethicalReading(true)))
                    {
                        PageCountDic.Add(i, PDF.GetNumberOfPages());
                    } // end using
                } // end for

                string TempAllCourseLearnResultPath = Path.Combine(DirectionPath, "TempCourseLearnResult.pdf");

                PdfHelper.MergeMultiple(CourseLearnResultFilePath, TempAllCourseLearnResultPath, false);

                //string CoverPDFFilePath = "合併PDF頁面/3-4 課程學習成果.pdf";

                string CoverPDFFilePath = Path.Combine(DirectionPath, "課程學習成果封面.pdf");

                string CoverHtmlString = HtmlHelper.GetHeader() + HtmlHelper.GenerateCourseLearnResultCoverHtmlString(CourseLearnResultList) + HtmlHelper.GetFooter();

                PdfHelper.RenderHtmlToPDF(CoverHtmlString, CoverPDFFilePath);

                int CoverPageCount = 0;

                using (PdfDocument CoverPDF = new PdfDocument(new PdfReader(CoverPDFFilePath)))
                {
                    CoverPageCount = CoverPDF.GetNumberOfPages();
                } // end using

                string TempCourseLearnResult = Path.Combine(DirectionPath, "課程學習成果1.pdf");

                PdfHelper.MergePDF(CoverPDFFilePath, TempAllCourseLearnResultPath, TempCourseLearnResult);

                if (File.Exists(TempAllCourseLearnResultPath))
                {
                    File.Delete(TempAllCourseLearnResultPath);
                } // end if

                if (File.Exists(CoverPDFFilePath))
                {
                    File.Delete(CoverPDFFilePath);
                } // end if

                string PDFFilePath = Path.Combine(DirectionPath, "課程學習成果.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempCourseLearnResult).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("課程學習成果", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetFirstPage()));

                    int PageCount = 0;

                    for (int i = 0; i < CourseLearnResultList.Count; i++)
                    {
                        PdfOutline SubOutline = PDFOutline.AddOutline(CourseLearnResultList[i]["CourseName"].ToString());

                        if (i == 0)
                        {
                            SubOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(CoverPageCount + 1)));
                        }
                        else
                        {
                            PageCount += PageCountDic[i - 1];

                            SubOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(PageCount + CoverPageCount + 1)));
                        } // end if

                        PDFOutline.GetAllChildren().Add(SubOutline);
                    } // end for
                } // end using

                if (File.Exists(TempCourseLearnResult))
                {
                    File.Delete(TempCourseLearnResult);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCourseLearnResult_RowData


        // 課程學習成果-PDF
        public static string GenerateCourseLearnResult_PDF(int EFID, string DirectionPath)
        {
            try
            {
                string CourseLearnResultFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string CoverFilePath = "合併PDF頁面/3-4 課程學習成果.pdf";

                string TempCourseRecordFilePath = Path.Combine(DirectionPath, "課程學習成果1.pdf");

                PdfHelper.MergePDF(CoverFilePath, CourseLearnResultFilePath, TempCourseRecordFilePath);

                string PDFFilePath = Path.Combine(DirectionPath, "課程學習成果.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempCourseRecordFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("課程學習成果", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempCourseRecordFilePath))
                {
                    File.Delete(TempCourseRecordFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCourseLearnResult_PDF


        // 專題實作-PDF
        public static string GenerateProfessionCourseLearnResult_PDF(int EFID, string DirectionPath)
        {
            try
            {
                string ProfessionCourseLearnResultFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string CoverFilePath = "合併PDF頁面/3-4-1專題實作及實習科目學習成果.pdf";

                string TempCourseRecordFilePath = Path.Combine(DirectionPath, "專題實作及實習科目學習成果1.pdf");

                PdfHelper.MergePDF(CoverFilePath, ProfessionCourseLearnResultFilePath, TempCourseRecordFilePath);

                string PDFFilePath = Path.Combine(DirectionPath, "專題實作及實習科目學習成果.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempCourseRecordFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("專題實作及實習科目學習成果", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempCourseRecordFilePath))
                {
                    File.Delete(TempCourseRecordFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateProfessionCourseLearnResult_PDF


        // 多元表現綜整心得
        public static string GenerateComprehensiveExperienceDoc(int EFID, string DirectionPath)
        {
            try
            {
                string CoverFilePath = "合併PDF頁面/3-5 多元表現綜整心得.pdf";

                string ProcessFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string TempPDFFilePath = Path.Combine(DirectionPath, "多元表現綜整心得1.pdf");

                string PDFFilePath = Path.Combine(DirectionPath, "多元表現綜整心得.pdf");

                PdfHelper.MergePDF(CoverFilePath, ProcessFilePath, TempPDFFilePath);

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempPDFFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("多元表現綜整心得", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempPDFFilePath))
                {
                    File.Delete(TempPDFFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateComprehensiveExperienceDoc


        // 多元表現
        public static string GenerateMultiplePerformance(JObject MultiplePerformance, string DirectionPath)
        {
            try
            {
                List<string> FilePathList = new List<string>();

                Dictionary<string, Dictionary<int, int>> PDFPageCountDic = new Dictionary<string, Dictionary<int, int>>();

                List<string> MultiplePerformanceItemName = new List<string>() { "SchoolCadre", "Competition", "Certification", "Cadre", "ServiceLearn", "AlternativeLearn", "JobLearn", "Collection", "TeamActivity", "OtherMultiplePerformance", "AdvancedPlacement" };

                foreach (string PropertyName in MultiplePerformanceItemName)
                {
                    JArray ItemList = JArray.Parse(MultiplePerformance[PropertyName].ToString());

                    if (ItemList.Count > 0)
                    {
                        Dictionary<int, int> ItemPageCountDic = new Dictionary<int, int>();

                        if (PropertyName == "SchoolCadre")
                        {
                            FilePathList.Add(GenerateSchoolCadre(ItemList, DirectionPath, out ItemPageCountDic));
                        }
                        else
                        {
                            FilePathList.Add(GenerateMultiplePerformanceItem(ItemList, DirectionPath, PropertyName, out ItemPageCountDic));
                        } // end if

                        PDFPageCountDic.Add(PropertyName, ItemPageCountDic);
                    } // end if
                } // end foreach

                string CoverFilePath = Path.Combine(DirectionPath, "多元表現封面.pdf");

                string MultiplePerformanceCoverHtmlString = HtmlHelper.GetHeader() + HtmlHelper.GenerateMultiplePerformanceCoverHtmlString(MultiplePerformance) + HtmlHelper.GetFooter();

                PdfHelper.RenderHtmlToPDF(MultiplePerformanceCoverHtmlString, CoverFilePath);

                int CoverPageCount = 0;

                using (PdfDocument CoverPDF = new PdfDocument(new PdfReader(CoverFilePath)))
                {
                    CoverPageCount = CoverPDF.GetNumberOfPages();
                } // end using

                string AllItemMergeFilePath = Path.Combine(DirectionPath, "TempAllMultiplePerformance.pdf");

                string AllItemWithCoverFilePath = Path.Combine(DirectionPath, "多元表現1.pdf");

                PdfHelper.MergeMultiple(FilePathList, AllItemMergeFilePath, false);

                PdfHelper.MergePDF(CoverFilePath, AllItemMergeFilePath, AllItemWithCoverFilePath, false);

                if (File.Exists(CoverFilePath))
                {
                    File.Delete(CoverFilePath);
                } // end if

                if (File.Exists(AllItemMergeFilePath))
                {
                    File.Delete(AllItemMergeFilePath);
                } // end if

                string PDFFilePath = Path.Combine(DirectionPath, "多元表現.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(AllItemWithCoverFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline Outline = PDF.GetOutlines(true);

                    Outline = Outline.AddOutline("多元表現", 0);

                    Outline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetFirstPage()));

                    int Index = CoverPageCount;

                    foreach (string PropertyName in MultiplePerformanceItemName)
                    {
                        JArray ItemList = JArray.Parse(MultiplePerformance[PropertyName].ToString());

                        if (ItemList.Count > 0)
                        {
                            if (PropertyName == "SchoolCadre")
                            {
                                Index += 1;

                                PdfOutline SubOutline = Outline.AddOutline(EPItemTask.GetMultiplePerformanceItemName(PropertyName));

                                SubOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(Index)));

                                Outline.GetAllChildren().Add(SubOutline);
                            }
                            else
                            {
                                PdfOutline ItemOutline = Outline.AddOutline(EPItemTask.GetMultiplePerformanceItemName(PropertyName));

                                ItemOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(Index + 1)));

                                for (int i = 0; i < ItemList.Count; i++)
                                {
                                    string TitleName = string.Empty;

                                    if (PropertyName == "Certification")
                                    {
                                        TitleName = JObject.Parse(ItemList[i]["LicenseInfo"].ToString())["證照名稱"].ToString();
                                    }
                                    else if (PropertyName == "JobLearn")
                                    {
                                        TitleName = $"{ItemList[i]["JobUnit"].ToString()}-{ItemList[i]["JobTitle"].ToString()}";
                                    }
                                    else if (PropertyName == "Cadre")
                                    {
                                        TitleName = $"{ItemList[i]["MajorName"].ToString()}-{ItemList[i]["PositionName"].ToString()}";
                                    }
                                    else
                                    {
                                        TitleName = ItemList[i]["MajorName"].ToString();
                                    } // end if

                                    PdfOutline ItemSubOutline = ItemOutline.AddOutline(TitleName);

                                    ItemSubOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(Index + 1)));

                                    Index += PDFPageCountDic[PropertyName][i];


                                    ItemOutline.GetAllChildren().Add(ItemSubOutline);
                                } // end for

                                Outline.GetAllChildren().Add(ItemOutline);
                            } // end if
                        } // end if                        
                    } // end foreach
                } // end using

                if (File.Exists(AllItemWithCoverFilePath))
                {
                    File.Delete(AllItemWithCoverFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateMultiplePerformance


        public static string GenerateSchoolCadre(JArray SchoolCadreList, string DirectionPath, out Dictionary<int, int> ItemPageCountDic)
        {
            try
            {
                string PDFFilePath = Path.Combine(DirectionPath, "校方建立幹部經歷紀錄.pdf");

                string Html = HtmlHelper.GetHeader() + HtmlHelper.GenerateSchoolCadreHtmlString(SchoolCadreList) + HtmlHelper.GetFooter();

                PdfHelper.RenderHtmlToPDF(Html, PDFFilePath);

                Dictionary<int, int> PageCountDic = new Dictionary<int, int>();

                PageCountDic.Add(0, 1);

                ItemPageCountDic = PageCountDic;

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateSchoolCadre


        public static string GenerateMultiplePerformanceItem(JArray ItemList, string DirectionPath, string ItemName, out Dictionary<int, int> ItemPageCountDic)
        {
            try
            {
                Func<JObject, string> HtmlFunc = EPItemTask.GetMultiplePerformanceToHtmlFunc(ItemName);

                List<string> FilePathList = new List<string>();

                Dictionary<int, int> PageCountDic = new Dictionary<int, int>();

                for (int i = 0; i < ItemList.Count; i++)
                {
                    string Html = HtmlHelper.GetHeader() + HtmlFunc.Invoke(JObject.Parse(ItemList[i].ToString())) + HtmlHelper.GetFooter();

                    string TempFilePath = Path.Combine(DirectionPath, $"{ItemName}{i}.pdf");

                    if (Path.GetExtension(ItemList[i]["DocUrl"].ToString()).ToLower() == ".pdf")
                    {
                        string ItemTempFilePath = Path.Combine(DirectionPath, $"{ItemName}Temp{i}.pdf");

                        PdfHelper.RenderHtmlToPDF(Html, ItemTempFilePath);

                        PdfHelper.MergePDF(ItemTempFilePath, Path.Combine(_DataParnetDirect, ItemList[i]["DocUrl"].ToString()), TempFilePath, false, true);

                        if (File.Exists(ItemTempFilePath))
                        {
                            File.Delete(ItemTempFilePath);
                        } // end if
                    }
                    else
                    {
                        PdfHelper.RenderHtmlToPDF(Html, TempFilePath);
                    } // end if

                    FilePathList.Add(TempFilePath);

                    using (PdfDocument PDF = new PdfDocument(new PdfReader(TempFilePath).SetUnethicalReading(true)))
                    {
                        PageCountDic.Add(i, PDF.GetNumberOfPages());
                    } // end using
                } // end for

                string ItemPDFFilePath = Path.Combine(DirectionPath, $"{ItemName}.pdf");

                PdfHelper.MergeMultiple(FilePathList, ItemPDFFilePath);

                ItemPageCountDic = PageCountDic;

                return ItemPDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCadre


        // 多元表現-PDF
        public static string GenerateMultiplePerformance_PDF(int EFID, string DirectionPath)
        {
            try
            {
                string MultiplePerformanceFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string CoverFilePath = "合併PDF頁面/3-6 多元表現.pdf";

                string TempMultiplePerformanceFilePath = Path.Combine(DirectionPath, "多元表現1.pdf");

                PdfHelper.MergePDF(CoverFilePath, MultiplePerformanceFilePath, TempMultiplePerformanceFilePath);

                string PDFFilePath = Path.Combine(DirectionPath, "多元表現.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempMultiplePerformanceFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("多元表現", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempMultiplePerformanceFilePath))
                {
                    File.Delete(TempMultiplePerformanceFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateMultiplePerformance_PDF


        // 其他文件表現-PDF
        public static string GenerateOtherDoc_PDF(int EFID, string DirectionPath)
        {
            try
            {
                string OtherDocFilePath = EPItemTask.GetStudentDocFilePathByEFID(EFID);

                string CoverFilePath = "合併PDF頁面/3-8 其他.pdf";

                string TempOtherDocFilePath = Path.Combine(DirectionPath, "其他1.pdf");

                PdfHelper.MergePDF(CoverFilePath, OtherDocFilePath, TempOtherDocFilePath);

                string PDFFilePath = Path.Combine(DirectionPath, "其他.pdf");

                using (PdfDocument PDF = new PdfDocument(new PdfReader(TempOtherDocFilePath).SetUnethicalReading(true), new PdfWriter(PDFFilePath)))
                {
                    PdfOutline PDFOutline = PDF.GetOutlines(false);

                    PDFOutline = PDFOutline.AddOutline("其他", -1);

                    PDFOutline.AddDestination(PdfExplicitDestination.CreateFit(PDF.GetPage(1)));
                } // end using

                if (File.Exists(TempOtherDocFilePath))
                {
                    File.Delete(TempOtherDocFilePath);
                } // end if

                return PDFFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateMultiplePerformance_PDF


        public static void SetMergePDFPath(int SID, string FilePath, int MergePDFEFID)
        {
            try
            {
                MySqlConnection Conn;

                MySqlTransaction Transaction = null;

                using (Conn = DBHelper.GetSqlConnection())
                {
                    try
                    {
                        if (Conn.State == System.Data.ConnectionState.Closed)
                        {
                            Conn.Open();
                        } // end if

                        Transaction = Conn.BeginTransaction();

                        MySqlCommand Command = Conn.CreateCommand();

                        Command.Transaction = Transaction;

                        if (MergePDFEFID <= 0)
                        {
                            Command.CommandText = "INSERT INTO experience_files (doc_url) VALUES (@FilePath);";

                            Command.Parameters.Add("@FilePath", MySqlDbType.VarChar).Value = FilePath;

                            int Result = Command.ExecuteNonQuery();

                            if (Result != 1)
                            {
                                Transaction.Rollback();

                                throw new Exception("更新資料庫失敗");
                            } // end if

                            long EFID = Command.LastInsertedId;

                            Command.CommandText = "UPDATE students_other_documents SET mergepdf_EFID = @EFID, mergepdf_file_type = 2 WHERE sid = @SID;";

                            Command.Parameters.Add("@EFID", MySqlDbType.Int32).Value = EFID;

                            Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                            Result = Command.ExecuteNonQuery();

                            if (Result != 1)
                            {
                                Transaction.Rollback();

                                throw new Exception("更新資料庫失敗");
                            } // end if
                        }
                        else
                        {
                            Command.CommandText = "UPDATE experience_files SET doc_url = @FilePath WHERE efid = @EFID;";

                            Command.Parameters.Add("@FilePath", MySqlDbType.VarChar).Value = FilePath;

                            Command.Parameters.Add("@EFID", MySqlDbType.Int32).Value = MergePDFEFID;

                            int Result = Command.ExecuteNonQuery();

                            if (Result != 1)
                            {
                                Transaction.Rollback();

                                throw new Exception("更新資料庫失敗");
                            } // end if
                        } // end if

                        Transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);
                    } // end try catch
                } // end using
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end SetMergePDFPath


        public static void SetMergeErrorEFID(int SID)
        {
            try
            {
                MySqlConnection Conn;

                MySqlTransaction Transaction = null;

                using (Conn = DBHelper.GetSqlConnection())
                {
                    try
                    {
                        if (Conn.State == System.Data.ConnectionState.Closed)
                        {
                            Conn.Open();
                        } // end if

                        Transaction = Conn.BeginTransaction();

                        MySqlCommand Command = Conn.CreateCommand();

                        Command.Transaction = Transaction;

                        Command.CommandText = "UPDATE students_other_documents SET mergepdf_EFID = @EFID, mergepdf_file_type = 0 WHERE sid = @SID;";

                        Command.Parameters.Add("@EFID", MySqlDbType.Int32).Value = -1;

                        Command.Parameters.Add("@SID", MySqlDbType.Int32).Value = SID;

                        int UpdateResult = Command.ExecuteNonQuery();

                        if (UpdateResult != 1)
                        {
                            Transaction.Rollback();

                            throw new Exception("更新資料庫失敗");
                        } // end if

                        Transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        Transaction.Rollback();

                        throw new Exception(ex.Message);
                    } // end try catch
                } // end using
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end SetMergeErrorEFID
    } // end StudentPDFHandle
}
