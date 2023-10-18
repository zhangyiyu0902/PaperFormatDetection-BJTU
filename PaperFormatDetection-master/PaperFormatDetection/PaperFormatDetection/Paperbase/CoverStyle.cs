﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using System.Xml;

namespace PaperFormatDetection.Paperbase
{
    class CoverStyle
    {
        /// <summary>
        /// 未避免变量过多，将同一部分变量放在一个数组里面 注意变量在数组中的顺序 
        /// 这里包含大标题，论文中文题目， 论文英文题目，学生信息，中文LOGO，英文LOGO五个数组
        /// 大标题9个参数 顺序分别是 0该部分是否必要 1标题之间空格数目 2对齐方式 3字体 4字体大小 5是否需要加粗 6段间距 7段前距 8段后距
        /// 论文中文题目8个参数 顺序分别是 0最大字数 1对齐方式 2字体 3是否需要加粗 4字体大小 5段间距 6段前距 7段后距
        /// 论文英文题目7个参数 顺序分别是 0居中方式 1字体 2是否需要加粗 3字体大小 4段间距 5段前距 6段后距
        /// 学生信息8个参数 顺序分别是 0最大字数 1字体 2居中方式 3字体大小 4段间距 5段前距 6段后距
        /// 中文LOGO7个参数 顺序分别是 0字体 1居中方式 2字体大小 3段间距 4段前距 5段后距
        /// 英文LOGO7个参数 顺序分别是 0字体 1居中方式 2字体大小 3段间距 4段前距 5段后距
        /// 若以后需要新增属性一定注意文件中各属性的顺序 因为为了方便 后面初始化赋值就按照这个顺序赋值 这相当于一个约定
        /// </summary>
        protected string[] coverstyleHeadline = new string[9];
        protected string[] coverstyleSubtitleCh = new string[8];
        protected string[] coverstyleSubtitleEn = new string[7];
        protected string[] coverstyleStuInformation = new string[7];
        protected string[] coverstyleSchoolNameCh = new string[6];
        protected string[] coverstyleDate = new string[6];

        public CoverStyle()
        {

        }
        //论文封面检测
        public void detectCoverStyle(List<Paragraph> list, WordprocessingDocument doc)
        {
            Tools.Util.printError("封面检测");
            Util.printError("----------------------------------------------");
            if (list.Count <= 0)
            {
                if (bool.Parse(coverstyleHeadline[0]))
                    Util.printError("论文缺少封面部分");
            }
            else
            {
                int YFlag = 0;
                int i = -1;
                //封面大标题
                while (YFlag == 0)
                {
                    i++;
                    if ((Util.getFullText(list[i]).Trim().Length != 0 && (Util.getFullText(list[i]).Replace(" ", "").IndexOf("本科") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("毕业") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("论文") != -1))) break;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                string CovHeadline = Util.getFullText(list[i]).Replace(" ", "");
                if (YFlag == 0)
                {
                    if (CovHeadline != "本科毕业设计（论文）")
                        Util.printError("论文类型" + "“" + CovHeadline + "”" + "错误，应为“本科毕业设计（论文）”");
                    if (!(Util.getFullText(list[i]).Trim().Length - Util.getFullText(list[i]).Trim().Replace(" ", "").Length == 0)) Util.printError("论文类型" + "“" + CovHeadline + "”" + "每个字之间不应包含空格");
                    if (!Util.correctJustification(list[i], doc, coverstyleHeadline[2])) Util.printError("论文类型" + "“" + CovHeadline + "”" + "对齐方式错误，应为" + coverstyleHeadline[2]);
                    if (!Util.correctSpacingBetweenLines_line(list[i], doc, coverstyleHeadline[6])) Util.printError("论文类型" + "“" + CovHeadline + "”" + "行间距错误，应为" + Util.DSmap[coverstyleHeadline[6]]);
                    if (!Util.correctSpacingBetweenLines_Be(list[i], doc, coverstyleHeadline[7])) Util.printError("论文类型" + "“" + CovHeadline + "”" + "段前间距错误，应为0行");
                    if (!Util.correctSpacingBetweenLines_Af(list[i], doc, coverstyleHeadline[8])) Util.printError("论文类型" + "“" + CovHeadline + "”" + "段后间距错误，应为0行");
                    if (!Util.correctfonts(list[i], doc, coverstyleHeadline[3], "Times New Roman")) Util.printError("论文类型" + "“" + CovHeadline + "”" + "字体错误，应为" + coverstyleHeadline[3]);
                    if (!Util.correctsize(list[i], doc, coverstyleHeadline[4])) Util.printError("论文类型" + "“" + CovHeadline + "”" + "字号错误，应为" + coverstyleHeadline[4]);
                    if (!Util.correctBold(list[i], doc, bool.Parse(coverstyleHeadline[5])))
                    {
                        if (bool.Parse(coverstyleHeadline[5]))
                            Util.printError("论文类型" + "“" + CovHeadline + "”" + "未加粗");
                        else
                            Util.printError("论文类型" + "“" + CovHeadline + "”" + "不需加粗");
                    }

                    //封面大标题中文括号检测
                    if (Util.getFullText(list[i]).IndexOf("(") != -1 || Util.getFullText(list[0]).IndexOf(")") != -1)
                        Util.printError("论文大标题的括号有误，应为中文括号");
                }

                //封面中文小标题
                while (YFlag == 0)
                {
                    i++;
                    if (Util.getFullText(list[i]).Trim().Length != 0) break;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                string SubChText = Util.getFullText(list[i]).Replace(" ", "").Trim();
                int SubChWordNum = Util.getFullText(list[i]).Replace(" ", "").Trim().Length;
                for (int j = 0; j < Util.getFullText(list[i]).Replace(" ", "").Trim().Length; j++)
                    if ((SubChText[j] >= 48 && SubChText[j] <= 57) || (SubChText[j] >= 65 && SubChText[j] <= 90) || (SubChText[j] >= 97 && SubChText[j] <= 122))
                    {
                        int k = 0;
                        int m = j;
                        for (; m < Util.getFullText(list[i]).Replace(" ", "").Trim().Length; m++)
                        {
                            if (!((SubChText[m] >= 48 && SubChText[m] <= 57) || (SubChText[m] >= 65 && SubChText[m] <= 90) || (SubChText[m] >= 97 && SubChText[m] <= 122))) break;
                            k++;
                        }
                        SubChWordNum = SubChWordNum - k + 1;
                        j = m;
                    }
                if (YFlag == 0)
                {
                    Console.WriteLine(Util.getFullText(list[i]));
                    if (SubChWordNum > int.Parse(coverstyleSubtitleCh[0])) Util.printError("论文中文标题字数超过20个");
                    if (!Util.correctJustification(list[i], doc, coverstyleSubtitleCh[1])) Util.printError("论文中文标题对齐方式错误，应为" + coverstyleSubtitleCh[1]);
                    if (!Util.correctSpacingBetweenLines_line(list[i], doc, coverstyleSubtitleCh[5])) Util.printError("论文中文标题行间距错误，应为" + Util.DSmap[coverstyleSubtitleCh[5]]);
                    if (!Util.correctSpacingBetweenLines_Be(list[i], doc, coverstyleSubtitleCh[6])) Util.printError("论文中文标题段前间距错误，应为0行");
                    if (!Util.correctSpacingBetweenLines_Af(list[i], doc, coverstyleSubtitleCh[7])) Util.printError("论文中文标题段后间距错误，应为0行");
                    if (!Util.correctfonts(list[i], doc, coverstyleSubtitleCh[2], "Times New Roman")) Util.printError("论文中文标题字体错误，应为" + coverstyleSubtitleCh[2]);
                    if (!Util.correctsize(list[i], doc, coverstyleSubtitleCh[4])) Util.printError("论文中文标题字号错误，应为" + coverstyleSubtitleCh[4]);
                    if (!Util.correctBold(list[i], doc, bool.Parse(coverstyleSubtitleCh[3])))
                    {
                        if (bool.Parse(coverstyleSubtitleCh[3]))
                            Util.printError("论文中文标题未加粗");
                        else
                            Util.printError("论文中文标题不需加粗");
                    }
                }

                //检测封面中文小标题之间是否有多余转行
                while (YFlag == 0)
                {
                    i++;
                    if (Util.getFullText(list[i]).Trim().Length != 0) break;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                string CNorEn_SubtitleText = Util.getFullText(list[i]).Replace(" ", "");
                if (!((CNorEn_SubtitleText[0] >= 65 && CNorEn_SubtitleText[0] <= 90) || (CNorEn_SubtitleText[0] >= 97 && CNorEn_SubtitleText[0] <= 122)))
                    Util.printError("封面中文小标题中间不应该有转行");

                //封面英文小标题
                string EnSubtitleString;
                int EnFlag = 0;
                while (YFlag == 0)
                {
                    EnSubtitleString = Util.getFullText(list[i]).Replace(" ", "");
                    if (Util.getFullText(list[i]).Trim().Length == 0) { EnFlag = 1; break; };
                    if ((EnSubtitleString[0] >= 65 && EnSubtitleString[0] <= 90) || (EnSubtitleString[0] >= 97 && EnSubtitleString[0] <= 122)) break;
                    else EnFlag = 1;
                    i++;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                if (YFlag == 0 && EnFlag == 0)
                {
                    if (!Util.correctJustification(list[i], doc, coverstyleSubtitleEn[0])) Util.printError("论文英文标题对齐方式错误，应为" + coverstyleSubtitleEn[0]);
                    //if (!Util.correctSpacingBetweenLines_line(list[i], doc, coverstyleSubtitleEn[4])) Util.printError("论文英文标题行间距错误，应为" + Util.DSmap[coverstyleSubtitleEn[4]]);
                    if (!Util.correctSpacingBetweenLines_Be(list[i], doc, coverstyleSubtitleEn[5])) Util.printError("论文英文标题段前间距错误，应为0行");
                    //if (!Util.correctSpacingBetweenLines_Af(list[i], doc, coverstyleSubtitleEn[6])) Util.printError("论文英文标题段后间距错误，应为0行");
                    if (!Util.correctfonts(list[i], doc, "Times New Roman", coverstyleSubtitleEn[1])) Util.printError("论文英文标题字体错误，应为" + coverstyleSubtitleEn[1]);
                    if (!Util.correctsize(list[i], doc, coverstyleSubtitleEn[3])) Util.printError("论文英文标题字号错误，应为" + coverstyleSubtitleEn[3]);
                    if (!Util.correctBold(list[i], doc, bool.Parse(coverstyleSubtitleEn[2])))
                    {
                        if (bool.Parse(coverstyleSubtitleEn[2]))
                            Util.printError("论文英文标题未加粗");
                        else
                            Util.printError("论文英文标题不需加粗");
                    }

                    //英文小标题字母大小写的检测
                    string[] CapitalEmptyWord = { "A", "Above", "An", "As", "Behind", "By", "But", "Before", "If", "The", "At", "For", "From", "Of", "Off", "On", "To", "In", "With", "And", "As", "While", "So" };
                    string[] SmallEmptyWord = { "a", "above", "an", "as", "behind", "by", "but", "before", "if", "the", "at", "for", "from", "of", "off", "on", "to", "in", "with", "and", "as", "while", "so" };
                    string SubtitleEnText = Util.getFullText(list[i]).Trim();
                    int SubtitleEnTextLength = Util.getFullText(list[i]).Trim().Length;
                    if (!(SubtitleEnText[0] >= 65 && SubtitleEnText[0] <= 90)) Util.printError("论文英文题目句首的单词首字母未大写");
                    int LastSpaceInText = SubtitleEnText.LastIndexOf(' ');
                    int NumInText = 0;
                    bool IsHaveEmptyWord = false;
                    for (; NumInText < SubtitleEnTextLength; NumInText++)
                    {
                        if (SubtitleEnText[NumInText] == ' ' && NumInText < LastSpaceInText)
                        {
                            string Word = "";
                            bool IsWordEmpty = true;
                            bool WordFlag = true;
                            int a = SubtitleEnText.IndexOf(' ', NumInText + 1);
                            Word = SubtitleEnText.Substring(NumInText + 1, a - NumInText - 1);
                            if (a - NumInText - 1 > 0)
                                IsWordEmpty = false;
                            else
                                IsHaveEmptyWord = true;
                            if ((Word[0] >= 65 && Word[0] <= 90) || (Word[0] >= 97 && Word[0] <= 122))
                            {
                                foreach (string S in CapitalEmptyWord)
                                {
                                    if (Word == S) Util.printError("论文英文题目非句首虚词首字母不应大写" + "  ----" + Word);
                                }
                                foreach (string C in SmallEmptyWord)
                                {
                                    if (Word == C) WordFlag = false;
                                }
                                if (WordFlag && !IsWordEmpty)
                                {
                                    if (!(Word[0] >= 65 && Word[0] <= 90) && Word.IndexOf("(") == -1) Util.printError("论文英文标题实词首字母未大写" + "  ----" + Word);
                                }
                            }
                        }
                        if (SubtitleEnText[NumInText] == ' ' && NumInText == LastSpaceInText)
                        {
                            bool WordFlag = true;
                            string Word = SubtitleEnText.Substring(NumInText + 1);
                            if ((Word[0] >= 65 && Word[0] <= 90) || (Word[0] >= 97 && Word[0] <= 122))
                            {
                                foreach (string S in CapitalEmptyWord)
                                {
                                    if (Word == S) Util.printError("论文英文题目非句首虚词首字母不应大写" + "  ----" + Word);
                                }
                                foreach (string C in SmallEmptyWord)
                                {
                                    if (Word == C) WordFlag = false;
                                }
                                if (WordFlag)
                                {
                                    if (Word[0] != '\0')
                                        if (!(Word[0] >= 65 && Word[0] <= 90) && Word.IndexOf("(") == -1) Util.printError("论文英文标题实词首字母未大写" + "  ----" + Word);
                                }
                            }
                        }
                    }
                    if (IsHaveEmptyWord) Util.printError("论文英文题目里单词间含有多余的空格");
                }

                //检测封面英文小标题之间是否有多余转行
                while (YFlag == 0)
                {
                    i++;
                    if (Util.getFullText(list[i]).Trim().Length != 0) break;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                //string SubOrStuInformationText = Util.getFullText(list[i]).Replace(" ", "");
                //if ((SubOrStuInformationText[0] >= 65 && SubOrStuInformationText[0] <= 90) || (SubOrStuInformationText[0] >= 97 && SubOrStuInformationText[0] <= 122))
                //    Util.printError("封面英文小标题中间不应该有转行");

                //封面学生信息
                int NumLost = 0;
                string StuInformationString;
                while (YFlag == 0)
                {
                    if ((Util.getFullText(list[i]).Trim().Length != 0 && (Util.getFullText(list[i]).Replace(" ", "").IndexOf("学院") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("专业") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("学生姓名") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("学号") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("指导教师") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("评阅教师") != -1 || Util.getFullText(list[i]).Replace(" ", "").IndexOf("完成日期") != -1))) break;
                    i++;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                int FirstStuInformation = i;
                int StuInformationLine = 0;
                string[] EachStuInformation = new string[] { "学院", "专业", "学生姓名", "学号", "指导教师" };
                if (YFlag == 0)
                {
                    for (; i <= FirstStuInformation + 4; i++)
                    {
                        StuInformationLine++;
                        StuInformationString = Util.getFullText(list[i]).Replace(" ", "");
                        if (Util.getFullText(list[i]).Trim().Length != 0)
                        {
                            if (StuInformationString.IndexOf(EachStuInformation[i - FirstStuInformation]) == -1)
                                Util.printError("第" + StuInformationLine + "行学生信息应为:" + EachStuInformation[i - FirstStuInformation]);
                            if (!Util.correctSpacingBetweenLines_line(list[i], doc, coverstyleStuInformation[4])) Util.printError("学生信息中" + "“" + EachStuInformation[i - FirstStuInformation] + "”" + "行行间距错误，应为" + Util.DSmap[coverstyleStuInformation[4]]);
                            if (!Util.correctSpacingBetweenLines_Be(list[i], doc, coverstyleStuInformation[5])) Util.printError("学生信息中" + "“" + EachStuInformation[i - FirstStuInformation] + "”" + "行学生信息段前间距错误，应为0行");
                            if (!Util.correctSpacingBetweenLines_Af(list[i], doc, coverstyleStuInformation[6])) Util.printError("学生信息中" + "“" + EachStuInformation[i - FirstStuInformation] + "”" + "行学生信息段后间距错误，应为0行");
                            if (!Util.correctfonts(list[i], doc, coverstyleStuInformation[1], "Times New Roman")) Util.printError("学生信息中" + "“" + EachStuInformation[i - FirstStuInformation] + "”" + "行学生信息字体错误，应为" + coverstyleStuInformation[1]);
                            if (!Util.correctsize(list[i], doc, coverstyleStuInformation[3])) Util.printError("学生信息中" + "“" + EachStuInformation[i - FirstStuInformation] + "”" + "行学生信息字号错误，应为" + coverstyleStuInformation[3]);
                        }
                        else NumLost++;
                    }
                    if (NumLost != 0) Util.printError("学生信息有" + NumLost + "行缺省");
                }

                //封面中文学校名
                while (YFlag == 0)
                {
                    i++;
                    if (Util.getFullText(list[i]).Trim().Length != 0) break;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                if (YFlag == 0)
                {
                    if (!Util.correctJustification(list[i], doc, coverstyleSchoolNameCh[1])) Util.printError("封面中文学校名对齐方式错误，应为" + coverstyleSchoolNameCh[1]);
                    if (!Util.correctSpacingBetweenLines_line(list[i], doc, coverstyleSchoolNameCh[3])) Util.printError("封面中文学校名行间距错误，应为" + Util.DSmap[coverstyleSchoolNameCh[3]]);
                    if (!Util.correctSpacingBetweenLines_Be(list[i], doc, coverstyleSchoolNameCh[4])) Util.printError("封面中文学校名段前间距错误，应为0行");
                    if (!Util.correctSpacingBetweenLines_Af(list[i], doc, coverstyleSchoolNameCh[5])) Util.printError("封面中文学校名段后间距错误，应为0行");
                    if (!Util.correctfonts(list[i], doc, coverstyleSchoolNameCh[0], "Times New Roman")) Util.printError("封面中文学校名字体错误，应为" + coverstyleSchoolNameCh[0]);
                    if (!Util.correctsize(list[i], doc, coverstyleSchoolNameCh[2])) Util.printError("封面中文学校名字号错误，应为" + coverstyleSchoolNameCh[2]);
                }


                //封面日期
                //封面英文学校名
                while (YFlag == 0)
                {
                    i++;
                    if (Util.getFullText(list[i]).Trim().Length != 0) break;
                    if (i >= list.Count) { YFlag = 1; break; };
                }
                if (YFlag == 0)
                {
                    if (!Util.correctJustification(list[i], doc, coverstyleDate[1])) Util.printError("封面日期对齐方式错误，应为" + coverstyleDate[1]);
                    if (!Util.correctSpacingBetweenLines_Be(list[i], doc, coverstyleDate[4])) Util.printError("封面日期段前间距错误，应为0行");
                    //if (!Util.correctSpacingBetweenLines_Af(list[i], doc, coverstyleDate[5])) Util.printError("封面英文学校名段后间距错误，应为0行");
                    if (!Util.correctfonts(list[i], doc, coverstyleDate[0], "Times New Roman")) Util.printError("封面日期字体错误，应为" + coverstyleDate[0]);
                    if (!Util.correctsize(list[i], doc, coverstyleDate[2])) Util.printError("封面日期字号错误，应为" + coverstyleDate[2]);
                }
            }
            Util.printError("----------------------------------------------");
        }
    }
}