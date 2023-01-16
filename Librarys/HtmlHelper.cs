using iText.Kernel.Geom;
using Microsoft.Extensions.Primitives;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MergeStudentPDF2.Librarys
{
    public class HtmlHelper
    {
        private static string _DataParentDirect = string.Empty;


        public static void SetDataParentDirect(string Direct)
        {
            _DataParentDirect = Direct;
        } // end SetDataParentDirect


        public static string GetHeader()
        {
            string Html = @"<!doctype html>
                                    <html>
                                        <head>
                                            <meta charset='UTF-8'>
                                            <title>@TitleName</title>
                                             <style>
                                                     body {
                                                            margin: 0 50px;
                                                            font-family: Microsoft JhengHei;
                                                            color:#000000;
                                                        }
                                                        @page{
                                                            margin: 0;
                                                        }
                                                        table {
                                                            width: 100%;
                                                            border-collapse: collapse;
                                                        }

                                                        .trDark {
                                                            background-color: #f0f0f0;
                                                        }

                                                        table tr:first-child td {
                                                            border-top: 2px solid #000;
                                                            border-bottom: 2px solid #000;
                                                        }

                                                        td {
                                                            color: #808080;
                                                            line-height: 30px;
                                                            padding: 0;
                                                        }

                                                        .tdName {
                                                            vertical-align: middle;
                                                            font-weight: bold;
                                                            width: 100px;
                                                            font-size: 26px;
                                                            color: #ce4d4d;
                                                        }

                                                        .tdYear {
                                                            font-weight: bold;
                                                            color: #ce4d4d;
                                                            width: 150px;
                                                        }

                                                        .tdHead {
                                                            background-color: #f0f0f0;
                                                            color: #808080;
                                                            font-weight: bold;
                                                        }

                                                        .tdRem {
                                                            font-family: 微軟正黑體;
                                                        }

                                                            .tdRem b {
                                                                font-family: MingLiU;
                                                            }

                                                        hr {
                                                            width: 100%;
                                                            height: 5px;
                                                            margin: 20px 0;
                                                            background-color: #c0c0c0;
                                                        } 

                                                        .tdData {
                                                            color: #000000;
                                                        }
                                                        tr td.tdData2:first-child:before {
                                                            content: '\0025FC   ';
                                                            color: #ce4d4d;
                                                        }

                                                        .tdCal {
                                                            border-top: 0px solid #000;
                                                            border-bottom: 0px solid #000;
                                                        }

                                                        .tdName2 {
                                                            vertical-align: top;
                                                            font-weight: bold;
                                                            width: 100px;
                                                            font-size: 20px;
                                                        }

                                                        .tdNameRec2 {
                                                            color: #ce4d4d;
                                                            font-size: 16px;
                                                        }

                                                        .tdHead2 {
                                                            -webkit-print-color-adjust: exact;
                                                            background-color: #f0f0f0;
                                                            color: #808080;
                                                            font-weight: bold;
                                                        }

                                                        .tdData2 {
                                                            padding: 5px;
                                                        }

                                                        .tdRem2 {
                                                            font-family: 微軟正黑體;
                                                        }
                                                        .tdRem2 b {
                                                            font-family: MingLiU;
                                                        }
			                                            .cBody {   
				                                            margin-top:60px;
				                                            margin-left:2%;
				                                            margin-right:2%;
				                                            margin-bottom: 90px;
				                                            color:#000000; 
			                                            }
			                                            .Title {
                                                            padding-top: 2em;
				                                            font-size : 2.5em;
				                                            text-align : center; 
				                                            border-style:solid;
				                                            border-color:#5B5B5B;
				                                            height:300px;
			                                            }
			                                            .Title div{
				                                            padding-top: 30px;
				                                            font-weight:bold; 
			                                            }
			                                            .DataBody{
				                                            padding-top : 100px;
				                                            width: 100%;
                                                            text-align:center;
			                                            }
			                                            .DataShow {
				                                            text-align:left; 
			                                            }
			                                            .TitleShow{
				                                            font-size: 1.5em;
				                                            padding-bottom: 30px;
				                                            width: 99%;
				
			                                            }
			                                            .DataBody .DataShow .DataTitle{
				                                            font-size:1.2em; 
				                                            float:left;
				                                            width:18%; 
				                                            border: 1px #cccccc solid; 
                                                            color:#000000;
				                                            padding: 2px;
                                                            /*background-color: #548dd4;*/
                                                            background-color: #ffffff;
															height: 35px;
															line-height: 35px;
			                                            }
			                                            .DataBody .DataShow .DataValue{ 
				                                            float:left;
				                                            width:18%;
				                                            padding:2px;
				                                            border: 1px #cccccc solid;
                                                            /*background-color: #dbe5f1;*/
                                                            background-color: #ffffff;
                                                            color:#000000;
															height: 35px;
															line-height: 35px;
			                                            }		
			                                            .DataBody .DataShow .DataTitle2{
				                                            font-size:1.2em; 
				                                            float:left;
				                                            width:22.5%;
                                                            /*background-color: #548dd4;*/
                                                            background-color: #ffffff;
				                                            border: 1px #cccccc solid;
                                                            color:#000000;
				                                            padding: 2px;
															height: 35px;
															line-height: 35px;
			                                            }
			                                            .DataBody .DataShow .DataValue2{ 
				                                            float:left;
				                                            width:22.5%;
				                                            padding: 2px;
				                                            border: 1px #cccccc solid;
                                                            /*background-color: #dbe5f1;*/
                                                            color:#000000;
															height: 35px;
															line-height: 35px;
			                                            }	
			                                            .DataBody .DataShow .DataTitle3{
				                                            font-size:1.2em; 
				                                            float:left;
				                                            width:47%;
                                                            /*background-color: #548dd4;*/
                                                            background-color: #ffffff;
				                                            border: 1px #cccccc solid;
                                                            color:#000000;
				                                            padding: 1.5px;
															height: 35px;
															line-height: 35px;
			                                            }
			                                            .DataBody .DataShow .DataValue3{ 
				                                            float:left;
				                                            width:47%;
				                                            padding: 1.5px;
				                                            border: 1px #cccccc solid;
                                                            /*background-color: #dbe5f1;*/
                                                            background-color: #ffffff;
                                                            color:#000000;
															height: 35px;
															line-height: 35px;
			                                            }	
			                                            .DataRow{
				                                            padding: 10px;
			                                            }
                                                        .cBody .baseinfo {
                                                            width: 100% !important;
                                                            border-collapse: collapse !important;
                                                        }
														
														.cBody .baseinfo tr:first-child td {
															border-top: 2px solid #cccccc !important;
															border-bottom: 2px solid #cccccc !important;
															font-size:1.2em !important;
														}
														
														.cBody .baseinfo td {
															font-size:1.2em !important;
														}
														
														.cBody .baseinfo tr td {
															border-top: 2px solid #cccccc !important;
															border-bottom: 2px solid #cccccc !important;
															font-size:1.2em !important;
														}
                                                </style>
                                        </head>
                                        <body>";
            return Html;
        } // end GetHeader


        public static string GetFooter()
        {
            string Html = @"</body> </html>";

            return Html;
        } // end GetFooter


        public static string GenerateBaseInfoHtmlString(JObject BaseInfo, decimal FiveSemesterAvgGrade, int CourseLearnCount, int MultiplePerformanceCount)
        {
            try
            {
                string FiveSemesterAvgGradeString = FiveSemesterAvgGrade == -1 ? "-" : FiveSemesterAvgGrade.ToString();

                string Html = GetHeader() + $@"<div class='cBody'>
                                    <img  src='data:image/png;base64,{Convert.ToBase64String(File.ReadAllBytes("合併PDF頁面/合併PDF-首頁封面.png"))}' width='100%' style='vertical-align: middle' />
                                    <div style='width:100%; height:30px'></div>
                                    <div style='background-color: #c00000; color: white; width:50%; font-size: 40px; text-align:center'>{BaseInfo["Name"]}</div>
                                    <div style='width:100%; height:30px'></div>
                                    <div style='background-image: linear-gradient(90deg, #ededed 85%, #c00000 15%); color: #ff4d4d; font-size: 34px; width:100%'><span style='margin-left: 50px; color: #c00000'>基本資料</span></div>
                                    <div style='width:100%; height:30px'></div>
                                    <table class='baseinfo' border='1'>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>編號/准考證號碼</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>學校名稱</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>學校地區</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>學習型態</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>五(或六)學期平均成績</span></td>
                                        </tr>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["RegNum"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["SchoolName"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["County"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["ExperimentalType"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{FiveSemesterAvgGradeString}</span></td>
                                        </tr>
                                    </table>
                                    <div class='DataRow' style='height: 10px'>&nbsp;</div>
                                    <table class='baseinfo' border='1'>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>群別</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>科班學程別</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>班別</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>部別</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>是否低收入戶</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>考生身份</span></td>
                                        </tr>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["GroupName"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["SubjectName"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["ClassName"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["InstituteName"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["LowIncome"]}</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{BaseInfo["CandStatus"]}</span></td>
                                        </tr>
                                    </table>
                                    <div class='DataRow' style='height: 10px'>&nbsp;</div>
                                    <table class='baseinfo' border='1'>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>項目</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>筆數</span></td>
                                        </tr>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>課程學習成果</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{CourseLearnCount}</span></td>
                                        </tr>
                                        <tr>
                                            <td style='text-align:left;'><span style='font-color:black !important'>多元表現(校方建立幹部經歷紀錄+自行勾選上傳多元表現)</span></td>
                                            <td style='text-align:left;'><span style='font-color:black !important'>{MultiplePerformanceCount}</span></td>
                                            
                                        </tr>
                                    </table>
                                     
                                 </div>";

                Html += GetFooter();

                return Html;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateBaseInfo


        public static string GenerateSemesterAverageGradeAndReletivePerformanceHtmlString(JArray SemesterAverageGradeList)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='9' align='center' style='text-align:center'>");
                SB.Append("            <div>學期平均成績及學習相對表現</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' rowspan='2' style='text-align:center;'>學年度</td>");
                SB.Append("        <td class='tdHead' rowspan='2' style='text-align:center;'>學期</td>");
                SB.Append("        <td class='tdHead' rowspan='2' style='text-align:center;'>學期平均成績</td>");
                SB.Append("        <td class='tdHead' colspan='3' style='text-align:center;border-bottom:1px solid #000;'>學習相對表現</td>");
                SB.Append("        <td class='tdHead' rowspan='2' style='text-align:center;'>學校名稱</td>");
                SB.Append("        <td class='tdHead' rowspan='2' style='text-align:center;'>群別</td>");
                SB.Append("        <td class='tdHead' rowspan='2' style='text-align:center;'>科班學程</td>");
                SB.Append("    </tr>");
                SB.Append("    <tr style='font-size:smaller;text-align:center;'>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>校</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>群</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>科班學程</td>");
                SB.Append("    </tr>");
                for (int i = 0; i < SemesterAverageGradeList.Count; i++)
                {
                    string Style = (i % 2 == 0) ? "" : "class=trDark";

                    JObject SemesterGradeObj = JObject.Parse(SemesterAverageGradeList[i].ToString());

                    SB.Append($"<tr {Style}>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["SchoolYear"]}</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["Semester"]}</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["SemesterGrade"]}</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["SchoolRelativePerformance"]}%</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["GroupRelativePerformance"]}%</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["SubjectRelativePerformance"]}%</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["SchoolGradeInfo"]["就讀學校資訊"]["名稱"]}</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["GroupGradeInfo"]["就讀群別資訊"]["名稱"]}</td>");
                    SB.Append($"    <td style='text-align:center;color:#000000;'>{SemesterGradeObj["SubjectGradeInfo"]["就讀科班學程別資訊"]["名稱"]}</td>");
                    SB.Append("</tr>");
                } // end for
                SB.Append("</table>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateFiveSemesterAverageGradeAndReletivePerformance


        public static string GenerateCourseRecordHtmlString(JObject CourseRecordObj)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                foreach (string propertyName in CourseRecordObj.Properties().Select(x => x.Name))
                {
                    SB.Append("<div style='width:100%;height: 30px'></div>");

                    SB.Append("<table>");
                    SB.Append("    <tr>");
                    SB.Append("        <td class='tdYear' colspan='7'  style='vertical-align:text-top;' align='center'>");
                    SB.Append($"            <div>{propertyName}</div>");
                    SB.Append("        </td>");
                    SB.Append("    </tr>");
                    SB.Append("    <tr>");
                    SB.Append("        <td style='width:60px;text-align:center;'>學年度</td>");
                    SB.Append("        <td style='width:60px;text-align:center;'>學期</td>");
                    SB.Append("        <td style='width:90px;text-align:center;'>課程類別</td>");
                    SB.Append("        <td style='text-align:center;'>科目名稱</td>");
                    SB.Append("        <td style='text-align:center;'>學分數</td>");
                    SB.Append("        <td style='width:90px;text-align:center;'>成績</td>");
                    SB.Append("        <td style='width:90px;text-align:center;'>推估相對表現</td>");
                    SB.Append("    </tr>");

                    JArray CourseRecoreList = JArray.Parse(CourseRecordObj[propertyName].ToString());

                    for (int i = 0; i < CourseRecoreList.Count; i++)
                    {
                        string Style = (i % 2 == 0) ? "" : "class=trDark";

                        string SchoolYear = CourseRecoreList[i]["SchoolYear"].ToString() == "-999" ? "-" : CourseRecoreList[i]["SchoolYear"].ToString();

                        string Semester = CourseRecoreList[i]["Semester"].ToString() == "-1" ? "-" : CourseRecoreList[i]["Semester"].ToString();

                        string CourseCategory = CourseRecoreList[i]["CourseCategory"].ToString() == string.Empty ? "-" : CourseRecoreList[i]["CourseCategory"].ToString();

                        string CourseCredits = CourseRecoreList[i]["CourseCredits"].ToString() == string.Empty ? "-" : CourseRecoreList[i]["CourseCredits"].ToString();

                        string Score = CourseRecoreList[i]["Score"].ToString() == string.Empty ? "-" : CourseRecoreList[i]["Score"].ToString();

                        string RelativePerformance = decimal.Parse(CourseRecoreList[i]["RelativePerformance"].ToString()) <= -1 ? "-" : CourseRecoreList[i]["RelativePerformance"].ToString();

                        SB.Append($"<tr {Style}>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{SchoolYear}</td>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{Semester}</td>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{CourseCategory}</td>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{CourseRecoreList[i]["CourseName"]}</td>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{CourseCredits}</td>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{Score}</td>");
                        SB.Append($"    <td class='tdData' style='text-align:center;color:#000000;'>{RelativePerformance}%</td>");
                        SB.Append("</tr>");
                    } // end for

                    SB.Append("</table>");

                    SB.Append("<div style='page-break-after: always'></div>");
                } // end foreach

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCourseRecordHtmlString


        public static string GenerateCourseLearnResultCoverHtmlString(JArray CourseLearnResultList)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width: 100%; height: 150px'></div>");
                SB.Append("<div style='text-align: center;' class='cBody'>");
                SB.Append(@$"<img  src='data:image/png;base64,{Convert.ToBase64String(File.ReadAllBytes("合併PDF頁面/課程學習成果.png"))}' width='100%' style='vertical-align: middle' />");
                SB.Append("<div style='width: 100%; height: 50px'></div>");
                SB.Append("<table class='baseinfo'>");
                SB.Append("    <tr>");
                SB.Append("        <td>學年度</td>");
                SB.Append("        <td>學期</td>");
                SB.Append("        <td>課程名稱</td>");
                SB.Append("        <td>成果類別</td>");
                SB.Append("    </tr>");
                for (int i = 0; i < CourseLearnResultList.Count; i++)
                {
                    string Style = i % 2 == 0 ? string.Empty : "class=trDark";

                    string SchoolYear = CourseLearnResultList[i]["SchoolYear"].ToString() == "-999" ? "-" : CourseLearnResultList[i]["SchoolYear"].ToString();

                    string Semester = CourseLearnResultList[i]["Semester"].ToString() == "-1" ? "-" : CourseLearnResultList[i]["Semester"].ToString();

                    string CourseName = CourseLearnResultList[i]["CourseName"].ToString() == string.Empty ? "-" : CourseLearnResultList[i]["CourseName"].ToString();

                    string AchievementType = CourseLearnResultList[i]["AchievementType"].ToString() == string.Empty ? "-" : (Convert.ToInt32(CourseLearnResultList[i]["AchievementType"].ToString()) == 1 ? "專題實作及實習科目學習成果" : "其他課程學習成果");

                    SB.Append($"<tr {Style}>");
                    SB.Append($"    <td>{SchoolYear}</td>");
                    SB.Append($"    <td>{Semester}</td>");
                    SB.Append($"    <td>{CourseName}</td>");
                    SB.Append($"    <td>{AchievementType}</td>");
                    SB.Append("</tr>");
                } // end for
                SB.Append("</table>");
                SB.Append("</div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCourseLearnResultCoverHtmlString


        public static string GenerateCourseLearnResultHtmlString(JObject CourseLearnResult)
        {
            try
            {
                StringBuilder SB = new StringBuilder();
                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='6' align='center'>");
                SB.Append("            <div>課程學習成果</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>科目名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>學年度</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>學期</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>成果類別</td>");
                //SB.Append("        <td class='tdHead' style='text-align:center;'>證明文件</td>");
                //SB.Append("        <td class='tdHead' style='text-align:center;'>證明文件影音檔</td>");
                SB.Append("    </tr>");

                string SchoolYear = CourseLearnResult["SchoolYear"].ToString() == "-999" ? "-" : CourseLearnResult["SchoolYear"].ToString();

                string Semester = CourseLearnResult["Semester"].ToString() == "-1" ? "-" : CourseLearnResult["Semester"].ToString();

                string CourseName = CourseLearnResult["CourseName"].ToString() == string.Empty ? "-" : CourseLearnResult["CourseName"].ToString();

                string AchievementType = CourseLearnResult["AchievementType"].ToString() == string.Empty ? "-" : (Convert.ToInt32(CourseLearnResult["AchievementType"].ToString()) == 1 ? "專題實作及實習科目學習成果" : "其他課程學習成果");

                string Desc = CourseLearnResult["Desc"].ToString() == string.Empty ? "-" : CourseLearnResult["Desc"].ToString();

                string DocUrl = CourseLearnResult["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = CourseLearnResult["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{CourseName}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{SchoolYear}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Semester}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AchievementType}</td>");
                //SB.Append($"    <td class='tdData' style='text-align:center;'>{DocUrl}</td>");
                //SB.Append($"    <td class='tdData' style='text-align:center;'>{VideoUrl}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='3' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                if (CourseLearnResult["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(CourseLearnResult["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != ".pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, CourseLearnResult["DocUrl"].ToString()));
                    } // end if
                } // end if

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCourseLearnResultHtmlString


        public static string GenerateImageHtmlString(string ExtensionName, string DocUrl)
        {
            try
            {
                string Base64 = Convert.ToBase64String(File.ReadAllBytes(System.IO.Path.Combine(_DataParentDirect, DocUrl)));

                return $"<img  src='data:image/{ExtensionName};base64,{Base64}' width='90%' />";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateImageHtmlString


        public static string GenerateMultiplePerformanceCoverHtmlString(JObject MultiplePerformance)
        {
            try
            {
                StringBuilder SB = new StringBuilder();
                SB.Append("<div style='width: 100%; height:50px'></div>");
                SB.Append(@$"<img  src='data:image/png;base64,{Convert.ToBase64String(File.ReadAllBytes("合併PDF頁面/多元表現.png"))}' width='100%' style='vertical-align: middle' />");
                SB.Append("<div class='cBody'>");
                SB.Append("<table class='baseinfo'>");
                SB.Append("    <tr>");
                SB.Append("       <td style='border-right: 2px solid #cccccc !important; border-left: 2px solid #cccccc !important;'>多元表現項目</td>");
                SB.Append("       <td style='border-right: 2px solid #cccccc !important'>名稱</td>");
                SB.Append("    </tr>");

                List<string> MultiplePerformanceItemName = new List<string>() { "SchoolCadre", "Competition", "Certification", "Cadre", "ServiceLearn", "AlternativeLearn", "JobLearn", "Collection", "TeamActivity", "OtherMultiplePerformance", "AdvancedPlacement" };

                foreach (string PropertyName in MultiplePerformanceItemName)
                {
                    JArray ItemList = JArray.Parse(MultiplePerformance[PropertyName].ToString());

                    if (ItemList.Count > 0)
                    {
                        string ItemName = EPItemTask.GetMultiplePerformanceItemName(PropertyName);

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
                            else if (PropertyName == "SchoolCadre" || PropertyName == "Cadre")
                            {
                                TitleName = ItemList[i]["MajorName"].ToString() + "-" + ItemList[i]["PositionName"].ToString();
                            }
                            else
                            {
                                TitleName = ItemList[i]["MajorName"].ToString();
                            } // end if

                            if (i == 0)
                            {
                                SB.Append("    <tr>");
                                SB.Append($"        <td rowspan='{ItemList.Count}' style='border-right: 2px solid #cccccc !important; border-left: 2px solid #cccccc !important;'>{ItemName}</td>");
                                SB.Append($"        <td style='border-right: 2px solid #cccccc !important'>{TitleName}</td>");
                                SB.Append("    </tr>");
                            }
                            else
                            {
                                SB.Append("    <tr>");
                                SB.Append($"        <td style='border-right: 2px solid #cccccc !important;'>{TitleName}</td>");
                                SB.Append("    </tr>");
                            } // end if
                        } // end for
                        
                    } // end if
                } // end foreach

                SB.Append("</table>");
                SB.Append("</div>");
                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch

        } // end GenerateMultiplePerformanceCoverHtmlString


        public static string GenerateSchoolCadreHtmlString(JArray SchoolCadreList)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='5' align='center'>");
                SB.Append("            <div>校方建立幹部經歷紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>單位名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>擔任職務名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>幹部等級</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>開始日期</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>結束日期</td>");
                SB.Append("    </tr>");

                for (int i = 0; i < SchoolCadreList.Count; i++)
                {
                    string Style = (i % 2 == 0) ? "" : "class=trDark";
                    SB.Append($"<tr {Style}>");
                    SB.Append($"    <td class='tdData' style='text-align:center;'>{SchoolCadreList[i]["MajorName"]}</td>");
                    SB.Append($"    <td class='tdData' style='text-align:center;'>{SchoolCadreList[i]["PositionName"]}</td>");
                    SB.Append($"    <td class='tdData' style='text-align:center;'>{EPItemTask.GetCadreLevelName(Convert.ToInt32(SchoolCadreList[i]["LevelID"].ToString()))}</td>");
                    SB.Append($"    <td class='tdData' style='text-align:center;'>{SchoolCadreList[i]["StartDate"]}</td>");
                    SB.Append($"    <td class='tdData' style='text-align:center;'>{SchoolCadreList[i]["EndDate"]}</td>");
                    SB.Append("</tr>");
                } // end for

                SB.Append("</table>");

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateSchoolCadretHtmlString


        public static string GenerateCadreHtmlString(JObject Cadre)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='7' align='center'>");
                SB.Append("            <div>幹部經歷暨事蹟紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>單位名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>擔任職務名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>幹部等級</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>開始日期</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>結束日期</td>");
                SB.Append("    </tr>");

                string Desc = Cadre["Desc"].ToString() == string.Empty ? "-" : Cadre["Desc"].ToString();

                string DocUrl = Cadre["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = Cadre["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = Cadre["VideoUrlExternal"].ToString() == string.Empty ? "-" : Cadre["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Cadre["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Cadre["PositionName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{EPItemTask.GetCadreLevelName(Convert.ToInt32(Cadre["LevelID"].ToString()))}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Cadre["StartDate"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Cadre["EndDate"]}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='4' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (Cadre["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(Cadre["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, Cadre["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateSchoolCadreHtmlString


        public static string GenerateCompeititionHtmlString(JObject Compition)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='7' align='center'>");
                SB.Append("            <div>競賽參與紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>競賽名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>競賽獎項</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>項目</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>競賽等級</td>");
                SB.Append("    </tr>");

                string Desc = Compition["Desc"].ToString() == string.Empty ? "-" : Compition["Desc"].ToString();

                string DocUrl = Compition["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = Compition["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = Compition["VideoUrlExternal"].ToString() == string.Empty ? "-" : Compition["VideoUrlExternal"].ToString();
                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Compition["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Compition["Awards"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Compition["Items"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{EPItemTask.GetCompetitionLevelName(Convert.ToInt32(Compition["LevelID"].ToString()))}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='3' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (Compition["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(Compition["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, Compition["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCompititionHtmlString


        public static string GenerateCertificationHtmlString(JObject Certification)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='6' align='center'>");
                SB.Append("            <div>檢定證照紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>證照名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>級數分數</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>結果分數</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>分項結果</td>");
                SB.Append("    </tr>");

                string CertificationName = Certification["LicenseInfo"] == null ? "-" : Certification["LicenseInfo"]["證照名稱"].ToString();

                string CertificationScore = Certification["LicenseInfo"] == null ? "-" : Certification["LicenseInfo"]["級數分數"].ToString();

                string Score = Convert.ToDecimal(Certification["Score"].ToString()) == -999.99M ? "-" : Certification["Score"].ToString();

                string Desc = Certification["Desc"].ToString() == string.Empty ? "-" : Certification["Desc"].ToString();

                string DocUrl = Certification["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = Certification["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = Certification["VideoUrlExternal"].ToString() == string.Empty ? "-" : Certification["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{CertificationName}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{CertificationScore}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Score}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Certification["ScoreList"]}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='3' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (Certification["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(Certification["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, Certification["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCertificationHtmlString


        public static string GenerateServiceLearnHtmlString(JObject ServiceLearn)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='7' align='center'>");
                SB.Append("            <div>服務學習紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>服務名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>單位名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>時數</td>");
                SB.Append("    </tr>");

                string ServiceHour = Convert.ToDecimal(ServiceLearn["ServiceHours"].ToString()) == -9999.999M ? "-" : ServiceLearn["ServiceHours"].ToString();

                string Desc = ServiceLearn["Desc"].ToString() == string.Empty ? "-" : ServiceLearn["Desc"].ToString();

                string DocUrl = ServiceLearn["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = ServiceLearn["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = ServiceLearn["VideoUrlExternal"].ToString() == string.Empty ? "-" : ServiceLearn["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{ServiceLearn["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{ServiceLearn["ServiceUnit"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{ServiceHour}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='2' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (ServiceLearn["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(ServiceLearn["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, ServiceLearn["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateServiceLearnHtmlString


        public static string GenerateAlternativeLearnHtmlString(JObject AlternativeLearn)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='8' align='center'>");
                SB.Append("            <div>彈性學習時間紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>內容(開設名稱)</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>開設單位</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>開設週數</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>每週節數</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>類別</td>");
                SB.Append("    </tr>");

                string Desc = AlternativeLearn["Desc"].ToString() == string.Empty ? "-" : AlternativeLearn["Desc"].ToString();

                string DocUrl = AlternativeLearn["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = AlternativeLearn["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = AlternativeLearn["VideoUrlExternal"].ToString() == string.Empty ? "-" : AlternativeLearn["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AlternativeLearn["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AlternativeLearn["AlternativeUnit"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AlternativeLearn["SessionWeeks"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AlternativeLearn["SessionPerWeek"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{EPItemTask.GetAlternativeLearnLevelName(Convert.ToInt32(AlternativeLearn["LevelID"].ToString()))}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='4' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (AlternativeLearn["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(AlternativeLearn["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, AlternativeLearn["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateAlternativeLearnHtmlString


        public static string GenerateJobLearnHtmlString(JObject JobLearn)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='8' align='center'>");
                SB.Append("            <div>職場學習紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>學習單位</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>職稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>時數</td>");
                SB.Append("    </tr>");

                string JobHours = Convert.ToDecimal(JobLearn["JobHours"].ToString()) == -9999.999M ? "-" : JobLearn["JobHours"].ToString();

                string Desc = JobLearn["Desc"].ToString() == string.Empty ? "-" : JobLearn["Desc"].ToString();

                string DocUrl = JobLearn["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = JobLearn["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = JobLearn["VideoUrlExternal"].ToString() == string.Empty ? "-" : JobLearn["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{JobLearn["JobUnit"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{JobLearn["JobTitle"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{JobHours}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='2' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrl}</div>");
                if (JobLearn["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(JobLearn["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, JobLearn["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateJobLearnHtmlString


        public static string GenerateCollectionHtmlString(JObject Collection)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='4' align='center'>");
                SB.Append("            <div>作品成果紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>作品日期</td>");
                SB.Append("    </tr>");

                string Desc = Collection["Desc"].ToString() == string.Empty ? "-" : Collection["Desc"].ToString();

                string DocUrl = Collection["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = Collection["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = Collection["VideoUrlExternal"].ToString() == string.Empty ? "-" : Collection["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Collection["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{Collection["CollectionDate"]}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='1' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (Collection["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(Collection["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, Collection["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateCollectionHtmlString


        public static string GenerateTeamActivityHtmlString(JObject TeamActivity)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='10' align='center'>");
                SB.Append("            <div>團體活動時間紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>內容名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>辦理單位</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>類別</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>節數或時數</td>");
                SB.Append("    </tr>");

                string ActivityHours = Convert.ToDecimal(TeamActivity["ActivityHours"].ToString()) == -9999.999M ? "-" : TeamActivity["ActivityHours"].ToString();

                string Desc = TeamActivity["Desc"].ToString() == string.Empty ? "-" : TeamActivity["Desc"].ToString();

                string DocUrl = TeamActivity["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = TeamActivity["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = TeamActivity["VideoUrlExternal"].ToString() == string.Empty ? "-" : TeamActivity["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{TeamActivity["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{TeamActivity["ActivityUnit"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{EPItemTask.GetTeamActivityLevelName(Convert.ToInt32(TeamActivity["LevelID"].ToString()))}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{ActivityHours}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='3' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (TeamActivity["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(TeamActivity["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, TeamActivity["DocUrl"].ToString()));
                    } // end if
                } // end if

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateTeamActivityHtmlString


        public static string GenerateOtherMultiplePerformanceHtmlString(JObject OtherMultiplePerformance)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='7' align='center'>");
                SB.Append("            <div>其他多元表現紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>主辦單位</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>時數</td>");
                SB.Append("    </tr>");

                string MultipleHours = Convert.ToDecimal(OtherMultiplePerformance["MultipleHours"].ToString()) == -9999.999M ? "-" : OtherMultiplePerformance["MultipleHours"].ToString();

                string Desc = OtherMultiplePerformance["Desc"].ToString() == string.Empty ? "-" : OtherMultiplePerformance["Desc"].ToString();

                string DocUrl = OtherMultiplePerformance["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = OtherMultiplePerformance["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = OtherMultiplePerformance["VideoUrlExternal"].ToString() == string.Empty ? "-" : OtherMultiplePerformance["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{OtherMultiplePerformance["MajorName"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{OtherMultiplePerformance["MultipleUnit"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{MultipleHours}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='2' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrlExternal}</div>");
                if (OtherMultiplePerformance["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(OtherMultiplePerformance["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, OtherMultiplePerformance["DocUrl"].ToString()));
                    } // end if
                } // end if

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateOtherMultiplePerformanceHtmlString


        public static string GenerateAdvancedPlacementHtmlString(JObject AdvancedPlacement)
        {
            try
            {
                StringBuilder SB = new StringBuilder();

                SB.Append("<div style='width:100%;height: 30px'></div>");

                SB.Append("<table>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdYear' colspan='9' align='center'>");
                SB.Append("            <div>大學及技專校院先修課程紀錄</div>");
                SB.Append("        </td>");
                SB.Append("    </tr>");
                SB.Append("    <tr>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>課程名稱</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>開設單位</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>學分數</td>");
                SB.Append("        <td class='tdHead' style='text-align:center;'>總時數</td>");
                SB.Append("    </tr>");

                string AdvancedHours = Convert.ToDecimal(AdvancedPlacement["AdvancedHours"].ToString()) == -9999.999M ? "-" : AdvancedPlacement["AdvancedHours"].ToString();

                string AdvancedCredits = Convert.ToDecimal(AdvancedPlacement["AdvancedCredits"].ToString()) == -1 ? "-" : AdvancedPlacement["AdvancedCredits"].ToString();

                string Desc = AdvancedPlacement["Desc"].ToString() == string.Empty ? "-" : AdvancedPlacement["Desc"].ToString();

                string DocUrl = AdvancedPlacement["DocUrl"].ToString() == string.Empty ? "-" : "請參閱下方檔案";

                string VideoUrl = AdvancedPlacement["VideoUrl"].ToString() == string.Empty ? "-" : "如要參閱請登入系統查看";

                string VideoUrlExternal = AdvancedPlacement["VideoUrlExternal"].ToString() == string.Empty ? "-" : AdvancedPlacement["VideoUrlExternal"].ToString();

                SB.Append("<tr>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AdvancedPlacement["AdvancedCourse"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AdvancedPlacement["AdvancedUnit"]}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AdvancedCredits}</td>");
                SB.Append($"    <td class='tdData' style='text-align:center;'>{AdvancedHours}</td>");
                SB.Append("</tr>");
                SB.Append("<tr>");
                SB.Append($"    <td style='font-weight:bold;width:180px;'>內容簡述：</td>");
                SB.Append($"    <td colspan='3' class='tdRem tdData' style='text-align:left;'>{Desc}</td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append($"<div>證明文件：{DocUrl}</div>");
                SB.Append($"<div>證明文件影音檔：{VideoUrl}</div>");
                SB.Append($"<div>外部影音連結：{VideoUrl}</div>");
                if (AdvancedPlacement["DocUrl"].ToString() != string.Empty)
                {
                    string ExtensionName = System.IO.Path.GetExtension(AdvancedPlacement["DocUrl"].ToString()).ToLower();

                    if (ExtensionName != "pdf")
                    {
                        SB.Append(GenerateImageHtmlString(ExtensionName, AdvancedPlacement["DocUrl"].ToString()));
                    } // end if
                } // end if                   

                SB.Append("<div style='page-break-after: always'></div>");

                return SB.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end GenerateAdvancedPlacementHtmlString
    } // end HtmlHelper
}
