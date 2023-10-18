using System;
using System.Xml;
using System.Collections.Generic;
using PaperFormatDetection.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace PaperFormatDetection.Undergraduate
{
    class Text : Paperbase.Text
    {
        public Text(WordprocessingDocument doc)
        {
            Util.printError("正文检测");
            Util.printError("----------------------------------------------");
            try
            {
                Init();
                detectAllText(sectionLoction(doc, "正文", 0), doc);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Util.printError("----------------------------------------------");
        }
        public void Init()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(Util.environmentDir + "/Template/Undergraduate/Text.xml");
            int m, index;
            XmlNodeList tTitleNode = xmlDoc.SelectSingleNode("Text").SelectSingleNode("tTitle").ChildNodes;
            index = 0;
            foreach (XmlNode levelnode in tTitleNode)
            {
                m = 0;
                XmlNodeList levelNode = levelnode.ChildNodes;
                foreach (XmlNode node in levelNode)
                {
                    tTitle[index, m] = node.InnerText;
                    m++;
                }
                index++;
            }

            XmlNodeList tTextNode = xmlDoc.SelectSingleNode("Text").SelectSingleNode("tText").ChildNodes;
            m = 0;
            foreach (XmlNode node in tTextNode)
            {
                tText[m] = node.InnerText;
                m++;
            }
        }
        public static List<DocumentFormat.OpenXml.OpenXmlElement> sectionLoction(WordprocessingDocument doc, string section, int paperType)
        {
            string[] Undergraduate = new string[] { "摘要", "Abstract", "目录", "引言", "正文", "结论", "参考文献", "附录", "致谢" };

            string[][] type = new string[][] { Undergraduate };
            int index = Array.IndexOf(type[paperType], section);
            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> eles = body.Elements();
            List<DocumentFormat.OpenXml.OpenXmlElement> elelist = new List<DocumentFormat.OpenXml.OpenXmlElement>();
            Boolean begin = false;
            Boolean end = false;
            if (section == "正文")
            {
                bool haveBookMark = false;

                foreach (DocumentFormat.OpenXml.OpenXmlElement p in eles)
                {
                    String fullText = "";
                    if (p.GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                        fullText = Util.getFullText((Paragraph)p).Trim();
                    if (p.GetFirstChild<BookmarkStart>() != null && p.GetFirstChild<BookmarkEnd>() != null)
                        haveBookMark = true;
                    else
                        haveBookMark = false;

                    if (fullText.Length > 0)
                    {
                        Match m = Regex.Match(fullText, @"[0-1]");
                        if (m.Success && m.Index == 0 && haveBookMark)
                        {
                            begin = true;
                        }
                    }
                    for (int i = index + 1; i < type[paperType].Length; i++)
                    {
                        if (fullText.Replace(" ", "").Length < 40 && fullText.Replace(" ", "").Equals(type[paperType][i]))
                        {
                            begin = false; end = true; break;
                        }
                    }
                    if (begin)
                    {
                        elelist.Add(p);
                    }
                    if (end)
                    {
                        break;
                    }
                }
                return elelist;
            }
            return elelist;
        }
    }
}
