using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace SatisfySurvey
{
    static class WordHelper
    {
        //学习情况调查前9题
        public const int A = 0, B = 1, C = 2;
        public static readonly int[] maxScore = new int[] 
        { 5, 5, 5, 5,10,10,
         10,10,10,10,10,10};

        public static int[][] satisfyCount = new int[9][];

        //学习情况调查第10题
        public static List<string> results10;

        //教师评分
        public static List<int>[] teacherCount = new List<int>[12];

        //满意度
        public static int[] final = new int[3];

        //建议
        public static List<string> suggests;

        //程序错误列表
        public static List<string> errorList = new List<string>();

        
        public static void InitHelper()
        {
            for (int i = 0; i < satisfyCount.Length; i++)
            {
                satisfyCount[i] = new int[4];
            }
            results10 = new List<string>();
            for (int i = 0; i < teacherCount.Length; i++)
            {
                teacherCount[i] = new List<int>();
            }
            suggests = new List<string>();
        }

 
        public static void DealWord(string name)
        {
            Application app;
            app = new ApplicationClass();

            object na = name;
            var doc = app.Documents.Open(ref na);
            try
            {
                DealWord(doc);
            }
            finally
            {
                doc.Close();
            }
        }

        public static void DealWord(Document doc)
        {

            int i = 1, j = 1;
            int k = 40;

            #region  满意度调查 前十题
            for (; i <= 60; i++)
            {
                string s = doc.Paragraphs[i].Range.Text.Replace("A.", "x").Replace("B.", "x").Replace("C.", "x");

                if (s.Contains(j + "."))
                {
                    
                    string answer = s.Split(new[] { j + "." },
                        StringSplitOptions.RemoveEmptyEntries)
                        [0].ToUpper();
                    if (answer.Contains('A'))
                    {
                        if (j == 7 && (s.Contains("ABC")))
                            satisfyCount[j - 1][C + 1]++;
                        else
                            satisfyCount[j - 1][A]++;
                    }
                    else if (answer.Contains('B'))
                        satisfyCount[j - 1][B]++;
                    else if (answer.Contains('C'))
                        satisfyCount[j - 1][C]++;
                    else if (j == 7 && answer.Contains('D'))
                        satisfyCount[j - 1][C + 1]++;
                    else if (j < 10)
                    {
                        throw new FormatException("[本题中没有提交答案: '" + s + "'] ");
                    }
                    j++;
                }

                if (s.Contains("你对于哪些知识点还存在疑问"))
                {
                    string sX = s;
                    for (int ii = 1; ii < 100; ii++)
                    {

                        if (doc.Paragraphs[i + ii].Range.Text.Contains("教 员 评 价 表"))
                        {
                            k = i + ii;
                            break;
                        }
                        sX += Environment.NewLine + doc.Paragraphs[i + ii].Range.Text;
                    }
                    string temp = sX.Replace("\n", "").Replace("\r", "").Replace(" ", "").Trim();

                    Console.WriteLine("sX.lengh = " + temp.Length);
                    if (temp.Length > 52)
                        results10.Add(sX.Trim('\n', '\r'));
                }
                if (s.Contains("评 价 项 目"))
                {
                    k = i;
                }

            }
            if (j < 11)
            {
                throw new FormatException("前十题 有误");
            }

            #endregion

            #region  满意度调查 教师评分
            for (i = k + 5, j = 1; i < k + 50; i += 4)
            {
                string s = doc.Paragraphs[i].Range.Text;
                s = s.Replace('\r', ' ').Replace('\n', ' ').Replace('\a', ' ').Trim();
                int _score = 100;
                try
                {
                    _score = ushort.Parse(s);
                }
                catch
                {
                    throw new FormatException("评分第" + j + "项 格式有误");
                }
                if (_score > maxScore[j - 1])
                    throw new ArgumentOutOfRangeException("评分第" + j+"项 评分超过限制");

                teacherCount[j - 1].Add(_score);
                j++;
            }

            #endregion
            #region  满意度调查 满意度
            k = i;

            for (i = i + 3, j = 0; j <= 3; i += 4, j++)
            {
                if (j == 3)
                    suggests.Add(Environment.NewLine + Environment.NewLine + "系统信息: " + doc.FullName + "没有提交“满意度”信息" + Environment.NewLine + Environment.NewLine);
                string s = doc.Paragraphs[i].Range.Text;
                // Console.WriteLine("line" + i + " " + s);
                s = s.Replace('\r', ' ').Replace('\n', ' ').Replace('\a', ' ').Trim();
                if (s != null && s.Length > 0)
                {
                    final[j]++;
                    break;
                }
            }
            #endregion
            #region  满意度调查 意见和建议
            i = k + 3 + 4 + 4 + 2;
            string suggest = "";
            for (; i <= doc.Paragraphs.Count; i++)
            {
                string s = doc.Paragraphs[i].Range.Text;
                suggest += s;
                Console.WriteLine("line" + i + " " + s);
            }
            suggest = suggest.Replace("意见和建议：", "").Replace("意见和建议", "").Trim(' ', '\n', '\r', '\a');
            if (suggest.Length > 1)
            {
                suggests.Add(suggest);

            }
            #endregion

            Console.WriteLine("over one word");
        }

        public static string Show()
        {
            string errors = "";
            string rn = Environment.NewLine, rn2 = Environment.NewLine + Environment.NewLine;
            if (errorList != null && errorList.Count != 0)
            {
                foreach (var err in errorList)
                {
                    errors += err + rn + rn;
                }
                return errors;
            }
            string output = errors;

            try
            {
                output += $"学习情况：{rn}";

                string studyinfo = File.ReadAllText("Inf.txt");
                for (int j = 0; j < 9; j++)
                {
                    for (int j2 = 0; j2 < 4; j2++)
                    {
                        if (satisfyCount[j][j2] != 0)
                            studyinfo = studyinfo.Replace("#" + j + j2 + "#", satisfyCount[j][j2] + "人");
                        else
                            studyinfo = studyinfo.Replace("#" + j + j2 + "#", "  ");
                    }
                }
                output += studyinfo + rn;
                for (int i = 0; i < results10.Count; i++)
                {
                    output += $"{"".PadRight(24, '*')}{rn}{results10[i]}{rn}{"".PadRight(24, '*')}";
                }

                int equal100 = 0, moreThan90 = 0, moreThan80 = 0, moreThan70 = 0, others = 0;
                for (int i = 0; i < teacherCount[0].Count; i++)
                {
                    int sum = 0;
                    for (int j = 0; j < teacherCount.Length; j++)
                    {
                        sum += teacherCount[j][i];
                    }
                    Console.WriteLine("sum = " + sum);
                    if (sum == 100)
                        equal100++;
                    else if (sum >= 90)
                        moreThan90++;
                    else if (sum >= 80)
                        moreThan80++;
                    else if (sum >= 70)
                        moreThan70++;
                    else
                        others++;

                }

                int all = equal100 + moreThan90 + moreThan80 + moreThan70 + others;
                float morethan90percent = (equal100 + moreThan90 + 0.0f) / all;
                float lessthan70percent = (others + 0.0f) / all;
                output += rn;
                output += $"分值统计：{rn}100分：{equal100}人   90分以上：{moreThan90}人   " +
                    $"80分以上：{moreThan80}人  70分以上：{moreThan70}人 " +
                        $"{rn2}百分比统计：{rn}> 90分：{morethan90percent * 100:N2} %    < 70分：{lessthan70percent * 100:N2} % {rn2}";


                output += $"{rn}满意度：{rn}";
                output += $"很满意： {final[0]}人{rn}"
                        + $"满意： {final[1]}人{rn}"
                        + $"不满意： {final[2]}人{rn2}";


                output += rn + "".PadRight(24, '*') + rn + "建议：" + rn2;

                for (int i = 0; i < suggests.Count; i++)
                {
                    output += rn + (i + 1) + ". " + suggests[i] + rn2;
                }

            }
            catch(Exception e)
            {
                output = "最后阶段处理异常 "+e.ToString();
            }
            return output;
        }
    }
}
