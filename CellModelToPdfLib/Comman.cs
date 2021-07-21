using CellModelToPdfLib.Model;
using Newtonsoft.Json;
using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CellModelToPdfLib
{
    public class Comman
    {
        internal static List<string> DeepCopyList(List<string> strList)
        {
            List<string> L = new List<string>();
            foreach (var p in strList)
            {
                L.Add(p);
            }
            return L;
        }

        public static XColor ToXColor(int color)
        {
            if (color == -1)
            {
                color = 1;
            }
            var t = ColorTranslator.FromWin32(color);
            var o = new XColor();
            o.A = 1;
            o.R = t.R;
            o.G = t.G;
            o.B = t.B;
            return o;
        }

        public static double mmmToPixels2(double v)
        {
            return (72 * (1 / 25.4f)) * (v / 10);
        }

        /// <summary>
        /// 字母转数字
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int StrTo26(string s)
        {
            int r = 0;
            for (int i = s.Length - 1; i >= 0; i--)
            {
                if (s[i] == '0')
                {
                    continue;
                }
                r += (Convert.ToInt32(s[i]) - 64) * Convert.ToInt32(Math.Pow(26, s.Length - 1 - i));
            }
            return r;
        }

        /// <summary>
        /// 数字转字母
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        public static string NumTo26(int num)
        {
            if (num <= 26)
            {
                return Chr(num + 64);
            }
            int t = num % 26;
            if (t == 0)
            {
                return Chr(num / 26 - 1 + 64) + "Z";
            }
            else
            {
                return Chr(num / 26 + 64) + Chr(t + 64);
            }

        }

        public static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }

        /// <summary>
        /// 拆分字串为词组
        /// </summary>
        /// <returns></returns>
        public static List<string> 拆分字串为词组(string s)
        {
            List<string> L = new List<string>() { "" };
            string s1 = "";
            int inQoute = 0;
            for (var i = 0; i < s.Length; i++)
            {
                if (s[i] == '"')
                {
                    if (i + 1 < s.Length && s[i + 1] == '"')
                    {
                        s1 += s[i];
                        i++;
                    }
                    else
                    {
                        inQoute++;
                    }
                    s1 += s[i];
                }
                else if (inQoute % 2 == 0)
                {
                    if (Regex.IsMatch(s[i].ToString(), @"\s"))
                    {
                        if (s1.Trim() != "")
                        {
                            L.Add(s1.Trim());
                            s1 = "";
                        }
                    }
                    else if ((s[i] == '>' || s[i] == '<') && s[i + 1] == '=')
                    {
                        if (s1.Trim() != "")
                        {
                            L.Add(s1.Trim());
                            s1 = "";
                        }
                        L.Add(s[i] + "=");
                        i += 1;
                    }
                    else if (s[i] == '<' && s[i + 1] == '>')
                    {
                        if (s1.Trim() != "")
                        {
                            L.Add(s1.Trim());
                            s1 = "";
                        }
                        L.Add("<>");
                        i += 1;
                    }
                    else if (s[i] == '(' || s[i] == ')' || s[i] == ',' || s[i] == '='
                        || s[i] == '>' || s[i] == '<' || s[i] == ';' || s[i] == '+'
                        || s[i] == '-' || s[i] == '*' || s[i] == '/' || s[i] == '%'
                        || s[i] == '!')
                    {
                        if (s1.Trim() != "")
                        {
                            L.Add(s1.Trim());
                            s1 = "";
                        }
                        L.Add(s[i].ToString());
                    }
                    else
                    {
                        s1 += s[i];
                    }
                }
                else
                {
                    s1 += s[i];
                }
            }
            if (s1.Trim() != "")
            {
                L.Add(s1.Trim());
            }
            return L;
        }

        public static List<T> DeepCopyList<T>(List<T> L)
        {
            string s = JsonConvert.SerializeObject(L);
            return JsonConvert.DeserializeObject<List<T>>(s);
        }

        public static bool IsNumber(string s)
        {
            if (Regex.IsMatch(s, @"^[+-]?\d+(.)?\d*$"))
            {
                return true;
            }
            return false;
        }

        public static void GetHVMargin(int fontsize, ref int hm, ref int hv)
        {
            //hm = Convert.ToInt32((fontsize / 10 + 3) / 2 + 1);
            //hv = Convert.ToInt32((fontsize / 10 + 2) / 2 + 1);            
            hm = 1;
            hv = 1;
        }

        public static bool IsColRowMark(string p)
        {
            if (Regex.IsMatch(p, @"^[a-zA-Z][a-zA-Z0-9]*?[0-9]$"))
            {
                return true;
            }
            return false;
        }

        public static void GetColRowFromStrMark(string strMark, ref int col, ref int row)
        {
            var i = 0;
            for (; i < strMark.Length; i++)
            {
                if (Regex.IsMatch(strMark[i].ToString(), @"[0-9]"))
                {
                    break;
                }
            }
            var part1 = strMark.Substring(0, i);
            var part2 = strMark.Substring(i);
            col = Comman.StrTo26(part1);
            row = Convert.ToInt32(part2);
        }

        public static void GetColRowFromStrMark(string strMark, ref string col, ref int row)
        {
            var i = 0;
            for (; i < strMark.Length; i++)
            {
                if (Regex.IsMatch(strMark[i].ToString(), @"[0-9]"))
                {
                    break;
                }
            }
            var part1 = strMark.Substring(0, i);
            var part2 = strMark.Substring(i);
            col = part1;
            row = Convert.ToInt32(part2);
        }

        public static CellInfo CopyCellInfo(CellInfo o)
        {
            var t = JsonConvert.SerializeObject(o);
            return JsonConvert.DeserializeObject<CellInfo>(t);
        }

        internal static double GetCellRangeHeight(CellJson cellJson, int sheet, int row1, int row2)
        {
            var t = 0d;
            for (var row = row1; row <= row2; row++)
            {
                t += cellJson.cells.Find(p => p.col == 1 && p.row == row && p.sheet == sheet).height;
            }
            return t;
        }
    }
}
