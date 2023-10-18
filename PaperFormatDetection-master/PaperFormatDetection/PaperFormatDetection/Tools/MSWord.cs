﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace PaperFormatDetection.Tools
{
    class MSWord
    {
        private Application wordApp = null;
        private Document wordDoc = null;
        private Dictionary<string, string> pageDic = new Dictionary<string, string>();//存储了所有表与图及其对应的页数
        public MSWord()
        {

        }
        public Dictionary<string, string> getPage(object filePath)
        {
            try
            {
                object missingValue = System.Reflection.Missing.Value;
                wordApp = new Application();
                wordApp.Visible = false;
                wordDoc = wordApp.Documents.Open(
                    ref filePath,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue);

                getTablesPage();
                getParagraphPage();

                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit();
                wordApp = null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                killWinWord();
            }
            return pageDic;
        }
        /// <summary>
        /// doc文档转为docx文档
        /// </summary>
        public string DocToDocx(object docPath)
        {
            string docxPath=docPath + "x";
            Application app = new Application();
            app.Visible = false;
            try
            {
                //如果已存在，则删除     
                if (File.Exists((string)docxPath))
                {
                    File.Delete((string)docxPath);
                }
                object missingValue = System.Reflection.Missing.Value;
                Document doc = app.Documents.Open(
                    ref docPath,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue);
                //object FileFormat = Word.WdSaveFormat.wdFormatDocument;
                object FileFormat = WdSaveFormat.wdFormatDocumentDefault;
                object LockComments = false;
                object Password = missingValue;
                object AddToRecentFiles = false;
                object WritePassword = missingValue;
                object ReadOnlyRecommended = false;
                object EmbedTrueTypeFonts = true;
                object SaveNativePictureFormat = missingValue;
                object SaveFormsData = missingValue;
                object SaveAsAOCELetter = missingValue;
                object Encoding = missingValue;
                object InsertLineBreaks = missingValue;
                object AllowSubstitutions = missingValue;
                object LineEnding = missingValue;
                object AddBiDiMarks = missingValue;
                object CompatibilityMode = missingValue;
                object docxFile = docxPath;
                doc.SaveAs(ref docxFile, ref FileFormat,
                    ref LockComments, ref Password, ref AddToRecentFiles, ref WritePassword,
                    ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat, ref SaveFormsData,
                    ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions,
                    ref LineEnding, ref AddBiDiMarks);
                //app.Documents.Close(ref missingValue, ref missingValue, ref missingValue);

                doc.Close();
                doc = null;
                app.Quit();
                app = null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                killWinWord();
            }
            finally
            {
            }
            return docxPath;
        }
        /// <summary>
        /// 杀掉winword.exe进程
        /// </summary>
        public void killWinWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }
        public void killWinWord()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }
        /// <summary>
        /// 把Word文件转换成pdf文件
        /// </summary>
        public bool WordToPdf(/*object sourcePath, string targetPath*/)
        {
            object docxPath=null;
            string pdfPath=null;
            
            bool result = false;
            WdExportFormat wdExportFormatPDF = WdExportFormat.wdExportFormatPDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Word.ApplicationClass applicationClass = null;
            Document document = null;
            try
            {
                applicationClass = new Microsoft.Office.Interop.Word.ApplicationClass();
                document = applicationClass.Documents.Open(ref docxPath, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                if (document != null)
                {
                    document.ExportAsFixedFormat(pdfPath, wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 0, 0, WdExportItem.wdExportDocumentContent, true, true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false, ref missing);
                }
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (document != null)
                {
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit(ref missing, ref missing, ref missing);
                    applicationClass = null;
                }
            }
            return result;
        }
       
        /// <summary>
        /// 获取文档中所有的表， 以/<第i张表， 页数/>的方式存储
        /// </summary>
        public void getTablesPage()
        {
            int NumberOfPage = -1;
            int count = 0;
            foreach (Table p in wordDoc.Tables)
            {
                count++;
                NumberOfPage = (int)p.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                //Console.WriteLine("2222   "+NumberOfPage);
                pageDic.Add("第"+count+"张表",NumberOfPage.ToString());
            }
        }
        public void getParagraphPage(/*string text*/)
        {
            List<string> chapterTitle=getHyperLinks();
            int NumberOfPage = -1;
            int prePage = -1;
            int nextPage = -1;
            int figurePage = -1;
            string Title = null;
            bool isMatch = false;
            int count = 0;
            Paragraph prePar = null;
            foreach (Paragraph p in wordDoc.Paragraphs)//每一段，先获取其内容
            {
                string text = p.Range.Text.Trim();
                //Util.printError(text);
                if (isMatch)
                {                   
                    nextPage = (int)p.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                    if (!pageDic.ContainsKey(Title))
                        pageDic.Add(Title, prePage.ToString() + "_" + NumberOfPage.ToString()+"_"+nextPage.ToString());
                    isMatch = false;
                }
               if (text.IndexOf('，')==-1 && text.Length>0  &&(Regex.IsMatch(text, @"^[表][ ]*?[0-9]") || chapterTitle.Contains(text) || Regex.IsMatch(text, @"^Fig")|| Regex.IsMatch(text, @"^Tab") || Regex.IsMatch(text, @"^[图][ ]*?[0-9]")))//如果读到的这个段落为标题或是图/表的标记
                {
                    Title = text;
                    NumberOfPage = (int)p.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                    prePage = (int)prePar.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                    //Console.WriteLine(NumberOfPage+"   "+text);
                    isMatch = true;
                }
                if (text != "")
                {
                    //判断该范围内是否存在图片
                    if (p.Range.InlineShapes.Count != 0)
                    {
                        foreach (InlineShape shape in p.Range.InlineShapes)
                        {
                            count++;
                            figurePage = (int)p.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                            //Console.WriteLine("   " + NumberOfPage);
                            pageDic.Add("第" + count + "张图", figurePage.ToString());
                        }
                        continue;
                    }                    
                }
                //if(text!="")
                    prePar = p;
            }
        }
        public List<string> getHyperLinks()//只有目录中标题会有超链接
        {
            List<string> chapterTitle = new List<string>();
            foreach (Hyperlink h in wordDoc.Hyperlinks)
            {
                string text = h.Range.Text.Trim();//获取超链接内容，山区首尾空格
                if (Regex.IsMatch(text, @"^[1-9][^.]"))//从字符串开头开始，前两个字符第一个为1~9， 第二个字符非.，因为目录里都是1.，2.类似
                {
                    text = Regex.Replace(text, @"[^\u4e00-\u9fa5_a-zA-Z]*$", "");//对于各种小标将其删除
                    chapterTitle.Add(text);
                }
            }
            return chapterTitle;
        }
    }
}

