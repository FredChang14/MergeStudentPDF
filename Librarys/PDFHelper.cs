
using iText.Forms;
using iText.Html2pdf;
using iText.Html2pdf.Resolver.Font;
using iText.IO.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Utils;
using iText.Layout;
using iText.Layout.Font;
using MySqlX.XDevAPI;
using Org.BouncyCastle.Asn1.X509;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static iText.Kernel.Pdf.PdfReader;

namespace MergeStudentPDF2.Librarys
{
    public class PdfHelper
    {
        private static List<string> FontList = new List<string>()
        {
            "Noto_Sans_TC/NotoSansTC-Black.otf",
            "Noto_Sans_TC/NotoSansTC-Bold.otf",
            "Noto_Sans_TC/NotoSansTC-Light.otf",
            "Noto_Sans_TC/NotoSansTC-Medium.otf",
            "Noto_Sans_TC/NotoSansTC-Regular.otf",
            "Noto_Sans_TC/NotoSansTC-Thin.otf"
        };


        public static void RenderHtmlToPDF(string Html, string FilePath)
        {
            try
            {
                ConverterProperties Properties = GetConverterProperties();

                PdfDocument PDF = new PdfDocument(new PdfWriter(FilePath));

                PDF.GetCatalog().SetPageMode(PdfName.UseOutlines);

                HtmlConverter.ConvertToPdf(Html, PDF, Properties);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end RenderHtmlToPDF


        public static void MergePDF(string FilePath1, string FilePath2, string OutputFilePath, bool MergeOutline = true, bool IsEPItemData = false)
        {
            try
            {
                if (IsEPItemData)
                {
                    using (PdfDocument PDF1 = new PdfDocument(new PdfReader(FilePath1).SetUnethicalReading(true)))
                    {
                        using (PdfDocument PDF2 = new PdfDocument(new PdfReader(FilePath2).SetUnethicalReading(true), new PdfWriter(OutputFilePath)))
                        {
                            int TotalPage = PDF2.GetNumberOfPages();

                            if (PDF2.HasOutlines())
                            {
                                PDF2.GetCatalog().Remove(PdfName.Outlines);
                            } // end if

                            for (int i = 1; i <= TotalPage; i++)
                            {
                                PdfPage Page = PDF2.GetPage(i);

                                IList<PdfAnnotation> PdfAnnotationList = Page.GetAnnotations();

                                if (PdfAnnotationList.Count == 0)
                                {
                                    continue;
                                } // end if

                                for (int j = 0; j < PdfAnnotationList.Count; j++)
                                {
                                    if (PdfAnnotationList[j].GetSubtype() == PdfName.Link)
                                    {
                                        Page.RemoveAnnotation(PdfAnnotationList[j]);
                                    } // end for
                                } // end for
                            } // end for

                            PDF1.CopyPagesTo(1, PDF1.GetNumberOfPages(), PDF2, 1);
                        } // end using
                    } // end using
                }
                else
                {
                    using (PdfDocument PDF1 = new PdfDocument(new PdfReader(FilePath1).SetUnethicalReading(true), new PdfWriter(OutputFilePath)))
                    {
                        using (PdfDocument PDF2 = new PdfDocument(new PdfReader(FilePath2).SetUnethicalReading(true)))
                        {
                            if (MergeOutline)
                            {
                                PDF1.GetCatalog().SetPageMode(PdfName.UseOutlines);
                            } // end if

                            int Page = PDF2.GetNumberOfPages();

                            PdfMerger PDFMerge = new PdfMerger(PDF1, false, MergeOutline);

                            PDFMerge.Merge(PDF2, 1, Page);

                            PDFMerge.Close();
                        } // end using
                    } // end using
                } // end if   
            }
            catch (Exception ex)
            {
                if (IsEPItemData)
                {
                    try
                    {
                        EPItemMergePDFExpection(FilePath1, FilePath2, OutputFilePath);
                    }
                    catch (Exception ex1)
                    {
                        throw new Exception(ex1.Message);
                    } // end try catch
                }
                else
                {
                    throw new Exception(ex.Message);
                } // end if
            } // end try catch
        } // end PdfDocument


        public static void MergeMultiple(List<string> FilePathList, string OutputFilePath, bool IsTopTitleMerge = true)
        {
            try
            {
                using (PdfDocument PDF1 = new PdfDocument(new PdfReader(FilePathList[0]).SetUnethicalReading(true), new PdfWriter(OutputFilePath)))
                {
                    PDF1.GetCatalog().SetPageMode(PdfName.UseOutlines);

                    PdfMerger PDFMerge = new PdfMerger(PDF1);

                    for (int i = 1; i < FilePathList.Count; i++)
                    {
                        using (PdfDocument PDF2 = new PdfDocument(new PdfReader(FilePathList[i]).SetUnethicalReading(true)))
                        {
                            PDFMerge.Merge(PDF2, 1, PDF2.GetNumberOfPages());
                        } // end using

                        if (File.Exists(FilePathList[i]))
                        {
                            File.Delete(FilePathList[i]);
                        } // end if
                    } // end for

                    PDFMerge.Close();
                } // end using

                if (File.Exists(FilePathList[0]))
                {
                    File.Delete(FilePathList[0]);
                } // end if
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end MergeMultiple


        public static void EPItemMergePDFExpection(string FilePath1, string FilePath2, string OutputFilePath)
        {
            try
            {
                using (PdfDocument OutputPDF = new PdfDocument(new PdfWriter(OutputFilePath)))
                {
                    using (PdfDocument PDF1 = new PdfDocument(new PdfReader(FilePath1).SetUnethicalReading(true)))
                    {
                        using (PdfDocument PDF2 = new PdfDocument(new PdfReader(FilePath2).SetUnethicalReading(true)))
                        {
                            PDF1.CopyPagesTo(1, PDF1.GetNumberOfPages(), OutputPDF);

                            PDF2.CopyPagesTo(1, PDF2.GetNumberOfPages(), OutputPDF);
                        } // end using
                    } // end using
                } // end using
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            } // end try catch
        } // end EPItemMergePDFExpection


        private static ConverterProperties GetConverterProperties()
        {
            ConverterProperties Properties = new ConverterProperties();

            FontProvider Provider = new DefaultFontProvider(false, false, false);

            foreach (string font in FontList)
            {
                FontProgram fontProgram = FontProgramFactory.CreateFont(font);

                Provider.AddFont(fontProgram);
            } // end foreach

            Properties.SetFontProvider(Provider);

            return Properties;
        } // end GetConverter
    } // end PDFHelper
}
