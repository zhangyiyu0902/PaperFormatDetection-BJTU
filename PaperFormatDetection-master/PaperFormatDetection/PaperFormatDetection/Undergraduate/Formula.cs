﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using PaperFormatDetection.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;

namespace PaperFormatDetection.Undergraduate
{
    class Formula : Paperbase.Formula
    {
        public Formula(WordprocessingDocument d)
        {
            Util.printError("公式检测");
            Util.printError("----------------------------------------------");
            try
            {
            Init();
            SelectandCheckFormula(d);
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
            xmlDoc.Load(Util.environmentDir + @"/Template/Undergraduate/Formula.xml");

            XmlNode justi = xmlDoc.SelectSingleNode("/Root/Formula/justification");//.SelectSingleNode("Justification");
            this.Justifi = justi.InnerText;

            XmlNode bef = xmlDoc.SelectSingleNode("/Root/Formula/linebefore");
            this.textBef = bef.InnerText;

            XmlNode aft = xmlDoc.SelectSingleNode("/Root/Formula/lineafter");
            this.textAft = aft.InnerText;

        }
    }
}
