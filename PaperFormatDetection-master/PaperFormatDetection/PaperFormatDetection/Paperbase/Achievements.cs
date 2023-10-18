using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using System.Text.RegularExpressions;

namespace PaperFormatDetection.Paperbase
{
    class Achievements
    {
        /// <summary>
        /// δ����������࣬��ͬһ���ֱ�������һ���������� ע������������е�˳�� 
        /// ����������۱��⣬�������ģ� ��л�������л�����ĸ�����
        /// ����9������ ˳��ֱ��� 0�ò����Ƿ��Ҫ 1����֮��ո���Ŀ 2���뷽ʽ 3�м�� 4��ǰ��� 5�κ��� 6���� 7�ֺ� 8�Ƿ�Ӵ�
        /// ����7������ ˳��ֱ��� 0���� 1���뷽ʽ 2�м�� 3��ǰ��� 4�κ��� 5���� 6�ֺ�
        /// ���Ժ���Ҫ��������һ��ע���ļ��и����Ե�˳�� ��ΪΪ�˷��� �����ʼ����ֵ�Ͱ������˳��ֵ ���൱��һ��Լ��
        /// </summary>
        protected string[] achievementTitle = new string[8];
        protected string[] achievementText = new string[7];
        protected string section = "";
        protected string beginList = "";

        public Achievements()
        {

        }
        //��ʿ�о��ɹ�  ˶ʿ���ķ���
        public void detectAchievement(List<Paragraph> list, WordprocessingDocument doc)
        {
            Util.printError(section + "���");
            Util.printError("----------------------------------------------");
            if (list.Count > 0)
            {
                //����
                if (!Util.correctJustification(list[0], doc, achievementTitle[1])) 
                    Util.printError(section + " ����δ" + achievementTitle[1]);
                if (!Util.correctSpacingBetweenLines_line(list[0], doc, achievementTitle[2]))
                    Util.printError(section + " �����м�����ӦΪ" + Util.DSmap[achievementTitle[2]]);
                if (!Util.correctSpacingBetweenLines_Be(list[0], doc, achievementTitle[3]))
                    Util.printError(section + " �����ǰ�����ӦΪ��ǰ0��");
                if (!Util.correctSpacingBetweenLines_Af(list[0], doc, achievementTitle[4]))
                    Util.printError(section + " ����κ�����ӦΪ�κ�1��");
                if (!Util.correctfonts(list[0], doc, achievementTitle[5], "Times New Roman")) 
                    Util.printError(section+" �����������ӦΪ"+achievementTitle[5]);
                if (!Util.correctsize(list[0], doc, achievementTitle[6])) 
                    Util.printError(section+" �����ֺŴ���ӦΪ"+achievementTitle[6]);
                if (!Util.correctBold(list[0], doc, bool.Parse(achievementTitle[7])))
                {
                    if (bool.Parse(achievementTitle[7]))
                        Util.printError(section+" ����δ�Ӵ�");
                    else
                        Util.printError(section+" ���ⲻ��Ӵ�");
                }
                //����
                bool isAchivement = false;
                List<Paragraph> achieList = new List<Paragraph>();
                int number = 1;//���Ķ�������
                string Rnumber = "^[0-9]";
                for (int i = 1; i < list.Count; i++)
                {
                    string temp = Tool.getFullText(list[i]).Trim();
                    string t = temp.Replace(" ", "");
                    string s = temp;
                    isAchivement = false;
                    if (temp.Length == 0) continue;
                    //if (temp.StartsWith(beginList)) isAchivement = true;
                    if (Regex.IsMatch(t, Rnumber)) isAchivement = true;
                    temp = "  ----" + (temp.Length > 10 ? temp.Substring(0, 10) : temp) + "......";
                    if (!Util.correctSpacingBetweenLines_line(list[i], doc, achievementText[1]))
                        Util.printError(section + " �м�����ӦΪ" + Util.DSmap[achievementText[1]] + temp);
                    if (!Util.correctSpacingBetweenLines_Be(list[i], doc, achievementText[2]))
                        Util.printError(section + " ��ǰ�����ӦΪ��ǰ0��" + temp);
                    if (!Util.correctSpacingBetweenLines_Af(list[i], doc, achievementText[3]))
                        Util.printError(section + " �κ�����ӦΪ�κ�0��" + temp);
                    if (!Util.correctfonts(list[i], doc, achievementText[4], "Times New Roman")) 
                        Util.printError(section + " �������ӦΪ"+achievementText[4] + temp);
                    if (!Util.correctsize(list[i], doc, achievementText[5])) 
                        Util.printError(section + " �ֺŴ���ӦΪ"+achievementText[5] + temp);
                    if(!isAchivement)
                    {
                        if (!Util.correctIndentation(list[i], doc, achievementText[0]))
                            Util.printError(section + " ��������ӦΪ��������" + achievementText[0] + "�ַ�" + temp);
                    }
                    else
                    {
                        //����ж�
                        int tnumber = Convert.ToInt32(t[0]) - 48;
                        if (tnumber != number)
                        {
                            //Util.printError(t[0].ToString());
                            Util.printError("��Ȩʹ����Ȩ�������Ŵ���ӦΪ" + number + " " + temp);

                        }
                        //��ź�Ŀո��ж�
                        if (s.Trim().Length > 3 && (s[1] != ' ' || s[2] != ' '))
                            Util.printError("��Ȩʹ����Ȩ��������������֮��Ӧ�������ո�" + " " + temp);
                        if (!Util.correctIndentation(list[i], doc, achievementText[6]))
                            Util.printError(section + " ��������ӦΪ��������" + achievementText[6] + "�ַ�" + temp);
                        achieList.Add(list[i]);
                        number++;
                    }
                }
                detectAlist(achieList);
            }
            Util.printError("----------------------------------------------");
        }
        public virtual void detectAlist(List<Paragraph> list)
        {

        }
        public bool containBold(Paragraph p, WordprocessingDocument doc)
        {
            bool containbold = false;
            string each = null;
            if (p != null)
            {
                IEnumerable<Run> rlist = p.Elements<Run>();
                foreach (Run run in rlist)
                {
                    each = Util.getFromRunPpr(run, 5); //��Run�����в���
                    if (each == null)
                    {
                        if (run.RunProperties != null)
                        {
                            RunStyle rs = run.RunProperties.RunStyle;
                            if (rs != null)
                            {
                                each = Util.getFromStyle(doc, rs.Val, 5);//��Runstyle�в���
                            }
                        }
                    }
                    if (each == null)
                    {
                        each = Util.getFromPpr(p, 5);//�Ӷ�����������
                        if (each == null && p.ParagraphProperties!=null)
                        {
                            ParagraphStyleId style_id = p.ParagraphProperties.ParagraphStyleId;
                            if (style_id != null)//��paragraphstyle�л�ȡ
                            {
                                each = Util.getFromStyle(doc, style_id.Val, 5);
                            }
                            if (each == null)//styleû�ҵ�
                            {
                                each = Util.getFromDefault(doc, 5);//�Ӷ���Ĭ��style����
                                if (each == null)//defaultû�ҵ�
                                {
                                    each = "false";
                                }
                            }
                        }
                    }
                    if (each == "true" && run.InnerText.ToString().Trim()!="")
                    {
                        containbold = true;
                    }
                }
            }
            return containbold;
        }
    }
}