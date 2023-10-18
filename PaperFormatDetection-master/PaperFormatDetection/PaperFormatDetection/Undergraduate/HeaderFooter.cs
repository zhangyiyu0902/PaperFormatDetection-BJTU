﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;
using PaperFormatDetection.Tools;

namespace PaperFormatDetection.Undergraduate
{
    class HeaderFooter : Paperbase.HeaderFooter
    {

        public HeaderFooter(WordprocessingDocument docx, string masterType)
        {
            Tools.Util.printError("页眉页脚检测");
            Tools.Util.printError("----------------------------------------------");
            try
            {
                Body body = docx.MainDocumentPart.Document.Body;
                Initiation(body);
                detectFirstSection(docx);
                detectUnderHeaderFooter(docx);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Tools.Util.printError("----------------------------------------------");

        }

        /*
        function:the first page and the second should have no header
        params: list:the list of sectPrs
        intlist:the location of sectPrs in body
        */
        protected void detectFirstSection(WordprocessingDocument docx)
        {
            Body body = docx.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            MainDocumentPart Mpart = docx.MainDocumentPart;
            List<SectionProperties> list = secPrList(body);
            SectionProperties s = null;

            if (list.Count == 0)
            {
                return;
            }
            s = list[0];


            IEnumerable<HeaderReference> headrs = s.Elements<HeaderReference>();
            IEnumerable<FooterReference> footrs = s.Elements<FooterReference>();
            foreach (HeaderReference headr in headrs)
            {
                string ID = headr.Id.ToString();
                HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID);
                Header header = hp.Header;
                if (header != null && header.InnerText != null)
                {
                    if (header.InnerText.Trim().Length > 0)
                    {
                        PaperFormatDetection.Tools.Util.printError("论文封面应无页眉");
                        break;
                    }
                }
            }
            foreach (FooterReference footr in footrs)
            {
                string ID = footr.Id.ToString();
                FooterPart fp = (FooterPart)Mpart.GetPartById(ID);
                Footer footer = fp.Footer;
                if (footer != null && footer.InnerText != null)
                {
                    if (footer.InnerText.Trim().Length > 0)
                    {
                        PaperFormatDetection.Tools.Util.printError("论文封面应无页脚");
                        break;
                    }
                }
            }
        }
        //本科论文页眉页脚检测
        private void detectUnderHeaderFooter(WordprocessingDocument docx)
        {
            Body body = docx.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            SectionProperties scetpr = body.GetFirstChild<SectionProperties>();
            MainDocumentPart Mpart = docx.MainDocumentPart;
            List<SectionProperties> list = secPrList(body);
            list.Add(scetpr);


            string docname = Util.getPaperName(docx);  //论文中文题目


            if (list.Count == 0)
                return;

            for (int i = 2; i <= list.Count<SectionProperties>(); i++)
            {

                SectionProperties s = list[i - 1];
                IEnumerable<HeaderReference> headrs = s.Elements<HeaderReference>();

                foreach (HeaderReference headr in headrs)
                {
                    #region 页眉
                    if (headr != null && headr.Type == HeaderFooterValues.Default)
                    {
                        string ID = headr.Id.ToString();
                        HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID);
                        Header header = hp.Header;

                        if (header.InnerText.Trim() != "")
                        {
                            Paragraph p = header.GetFirstChild<Paragraph>();
                            string headername = header.InnerText.Trim();
                            if (Util.correctfonts(p, docx, CNheaderFonts, ENHeaderFonts) == false)
                            {
                                PaperFormatDetection.Tools.Util.printError("第" + i + "节" + "页眉字体错误，应为:" + CNheaderFonts);
                            }

                            if (Util.correctsize(p, docx, headerFontsize) == false)
                            {
                                PaperFormatDetection.Tools.Util.printError("第" + i + "节" + "页眉字号错误，应为：" + headerFontsize);
                            }
                            if (Util.correctJustification(p, docx, headerJustification) == false)
                            {
                                PaperFormatDetection.Tools.Util.printError("第" + i + "节" + "页眉未居中" );
                            }
                            if (headername != docname)
                            {
                                PaperFormatDetection.Tools.Util.printError("第" + i + "节" + "页眉内容错误，应为：" + docname);
                            }
                        }
                    }
                    #endregion
                }

                IEnumerable<FooterReference> footrs = s.Elements<FooterReference>();
                foreach (FooterReference footr in footrs)
                {

                    if (footr != null)
                    {
                        string ID = footr.Id.ToString();
                        FooterPart fp = (FooterPart)Mpart.GetPartById(ID);
                        Footer footer = fp.Footer;
                        Paragraph p = footer.GetFirstChild<Paragraph>();

                        if (footer.InnerText != null && footer.InnerText.Trim() != "")
                        {
                            if (Util.correctfonts(p, docx, CNFooterFonts, ENFooterFonts) == false)
                            {
                                Util.printError("第" + i + "节" + "页脚字体错误，应为：" + CNFooterFonts);
                            }

                            if (Util.correctsize(p, docx, footerFontsize) == false)
                            {
                                PaperFormatDetection.Tools.Util.printError("第" + i + "节" + "页脚字号错误，应为：" + footerFontsize);
                            }
                            if (Util.correctJustification(p, docx, footerJustification) == false)
                            {
                                PaperFormatDetection.Tools.Util.printError("第" + i + "节" + "页脚未居中");
                            }
                        }
                    }

                }

            }



        }

        //从模板文件中读取数据
        public void Initiation(Body body)
        {
            string modelPath = Util.environmentDir + "\\Template\\Undergraduate\\HeaderFooter.xml";
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            try
            {
                XmlDocument myXmlDoc = new XmlDocument();
                myXmlDoc.Load(modelPath);
                XmlNode rootNode = myXmlDoc.SelectSingleNode("HeaderFooter");

                XmlNode HeaderNode = rootNode.SelectSingleNode("Header");
                this.CNheaderFonts = HeaderNode.SelectSingleNode("CNheaderFonts").InnerText;
                this.headerFontsize = HeaderNode.SelectSingleNode("headerFontsize").InnerText;
                this.headerJustification = HeaderNode.SelectSingleNode("headerJustification").InnerText;

                XmlNode FooterNode = rootNode.SelectSingleNode("Footer");
                this.CNFooterFonts = FooterNode.SelectSingleNode("CNFooterFonts").InnerText;
                this.footerFontsize = FooterNode.SelectSingleNode("footerFontsize").InnerText;
                this.footerJustification = FooterNode.SelectSingleNode("footerJustification").InnerText;

                //this.indexOfChineseTitleOfDocx = int.Parse(rootNode.SelectSingleNode("indexOfChineseTitleOfDocx").InnerText);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }





}
