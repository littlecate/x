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

        internal static double GetCellRangeWidth(CellJson cellJson, int sheet, int col1, int col2)
        {
            var t = 0d;
            for (var col = col1; col <= col2; col++)
            {
                t += cellJson.cells.Find(p => p.row == 1 && p.col == col && p.sheet == sheet).width;
            }
            return t;
        }

        internal static void 将某行以后的行整体移动(CellJson cellJson, int row2, double addHeight, int sheet)
        {
            List<string> markHaveDisposedList = new List<string>();
            for (var k = 1; k < cellJson.cols; k++)
            {
                for (var i = row2 + 1; i < cellJson.rows; i++)
                {
                    var o1 = cellJson.cells.Find(p => p.col == k && p.row == i && p.sheet == sheet);
                    o1.y += addHeight;
                    if (o1.cellPostionList != null)
                    {
                        var key1 = k + "_" + i + "_" + i;
                        if (!markHaveDisposedList.Contains(key1))
                        {
                            var L5 = o1.cellPostionList.FindAll(p => p.row == i);
                            for (var n = 0; n < L5.Count; n++)
                            {
                                L5[n].y += addHeight;
                            }
                            markHaveDisposedList.Add(key1);
                        }
                    }
                    if (o1.cellPositionTotal != null)
                    {
                        o1.cellPositionTotal.y += addHeight;
                    }
                    if (o1.mergeTo != null)
                    {
                        var c = o1.mergeTo.col;
                        var r = o1.mergeTo.row;
                        if (c != 0 && r != 0)
                        {
                            var oo = cellJson.cells.Find(p => p.col == c && p.row == r && p.sheet == sheet);
                            if (oo != null)
                            {
                                var key = c + "_" + r + "_" + i;
                                if (!markHaveDisposedList.Contains(key))
                                {
                                    var L5 = oo.cellPostionList.FindAll(p => p.row == i);
                                    for (var n = 0; n < L5.Count; n++)
                                    {
                                        L5[n].y += addHeight;
                                    }
                                    markHaveDisposedList.Add(key);
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static void 将某列以后的列整体移动(CellJson cellJson, int col2, double addWidth, int sheet)
        {
            List<string> markHaveDisposedList = new List<string>();
            for (var k = 1; k < cellJson.rows; k++)
            {
                for (var i = col2 + 1; i < cellJson.cols; i++)
                {
                    var o1 = cellJson.cells.Find(p => p.row == k && p.col == i && p.sheet == sheet);
                    o1.x += addWidth;
                    if (o1.cellPostionList != null)
                    {
                        var key1 = k + "_" + i + "_" + i;
                        if (!markHaveDisposedList.Contains(key1))
                        {
                            var L5 = o1.cellPostionList.FindAll(p => p.col == i);
                            for (var n = 0; n < L5.Count; n++)
                            {
                                L5[n].x += addWidth;
                            }
                            markHaveDisposedList.Add(key1);
                        }
                    }
                    if (o1.cellPositionTotal != null)
                    {
                        o1.cellPositionTotal.x += addWidth;
                    }
                    if (o1.mergeTo != null)
                    {
                        var c = o1.mergeTo.col;
                        var r = o1.mergeTo.row;
                        if (c != 0 && r != 0)
                        {
                            var oo = cellJson.cells.Find(p => p.col == c && p.row == r && p.sheet == sheet);
                            if (oo != null)
                            {
                                var key = c + "_" + r + "_" + i;
                                if (!markHaveDisposedList.Contains(key))
                                {
                                    var L5 = oo.cellPostionList.FindAll(p => p.col == i);
                                    for (var n = 0; n < L5.Count; n++)
                                    {
                                        L5[n].x += addWidth;
                                    }
                                    markHaveDisposedList.Add(key);
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static void 设置某行到某行的整体行高(CellJson cellJson, int row1, int row2, int sheet)
        {
            List<string> markHaveDisposedList = new List<string>();
            for (var i = 1; i < cellJson.cols; i++)
            {
                for (var k = row1; k <= row2; k++)
                {
                    var o = cellJson.cells.Find(p => p.col == i && p.row == k && p.sheet == sheet);
                    if (o.cellPositionTotal != null)
                    {
                        o.cellPositionTotal.width = o.cellPostionList.FindAll(p => p.row == k).Sum<CellPosition>(p => p.width);
                        o.cellPositionTotal.height = o.cellPostionList.FindAll(p => p.col == i).Sum<CellPosition>(p => p.height);
                    }
                    if (o.mergeTo != null)
                    {
                        var c = o.mergeTo.col;
                        var r = o.mergeTo.row;
                        if (c != 0 && r != 0)
                        {
                            var key = c + "_" + r;
                            if (!markHaveDisposedList.Contains(key))
                            {
                                var oo = cellJson.cells.Find(p => p.col == c && p.row == r && p.sheet == sheet);
                                if (oo != null)
                                {
                                    oo.cellPositionTotal.width = oo.cellPostionList.FindAll(p => p.row == k).Sum<CellPosition>(p => p.width);
                                    oo.cellPositionTotal.height = oo.cellPostionList.FindAll(p => p.col == i).Sum<CellPosition>(p => p.height);
                                }
                                markHaveDisposedList.Add(key);
                            }
                        }
                    }
                }
            }
        }

        internal static void 设置某列到某列的整体列宽(CellJson cellJson, int col1, int col2, int sheet)
        {
            List<string> markHaveDisposedList = new List<string>();
            for (var i = 1; i < cellJson.rows; i++)
            {
                for (var k = col1; k <= col2; k++)
                {
                    var o = cellJson.cells.Find(p => p.row == i && p.col == k && p.sheet == sheet);
                    if (o.cellPositionTotal != null)
                    {
                        o.cellPositionTotal.width = o.cellPostionList.FindAll(p => p.row == i).Sum<CellPosition>(p => p.width);
                        o.cellPositionTotal.height = o.cellPostionList.FindAll(p => p.col == k).Sum<CellPosition>(p => p.height);
                    }
                    if (o.mergeTo != null)
                    {
                        var c = o.mergeTo.col;
                        var r = o.mergeTo.row;
                        if (c != 0 && r != 0)
                        {
                            var key = c + "_" + r;
                            if (!markHaveDisposedList.Contains(key))
                            {
                                var oo = cellJson.cells.Find(p => p.col == c && p.row == r && p.sheet == sheet);
                                if (oo != null)
                                {
                                    oo.cellPositionTotal.width = oo.cellPostionList.FindAll(p => p.row == i).Sum<CellPosition>(p => p.width);
                                    oo.cellPositionTotal.height = oo.cellPostionList.FindAll(p => p.col == k).Sum<CellPosition>(p => p.height);
                                }
                                markHaveDisposedList.Add(key);
                            }
                        }
                    }
                }
            }
        }

        internal static void 以某区域为参考对齐列和列宽(CellJson cellJson, int col1, int col2, int sheet, List<CellInfo> L2)
        {
            List<string> markHaveDisposedList = new List<string>();
            for (var i = col1; i <= col2; i++)
            {
                var o = L2.Find(p => p.col == i);
                var x = o.x;
                var width = o.width;
                var L4 = cellJson.cells.FindAll(p => p.col == i && p.sheet == sheet);
                for (var k = 0; k < L4.Count; k++)
                {
                    var o1 = L4[k];
                    o1.x = x;
                    o1.width = width;
                    if (o1.cellPostionList != null)
                    {
                        var key1 = i + "_" + k + "_" + i;
                        if (!markHaveDisposedList.Contains(key1))
                        {
                            var L5 = o1.cellPostionList.FindAll(p => p.col == i);
                            for (var n = 0; n < L5.Count; n++)
                            {
                                L5[n].x = x;
                                L5[n].width = width;
                            }
                            markHaveDisposedList.Add(key1);
                        }
                    }
                    if (o1.cellPositionTotal != null)
                    {
                        o1.cellPositionTotal.x = x;
                    }
                    if (o1.mergeTo != null)
                    {
                        var c = o1.mergeTo.col;
                        var r = o1.mergeTo.row;
                        if (c != 0 && r != 0)
                        {
                            var oo = cellJson.cells.Find(p => p.col == c && p.row == r && p.sheet == sheet);
                            if (oo != null)
                            {
                                var key = c + "_" + r + "_" + i;
                                if (!markHaveDisposedList.Contains(key))
                                {
                                    var L5 = oo.cellPostionList.FindAll(p => p.col == i);
                                    for (var n = 0; n < L5.Count; n++)
                                    {
                                        L5[n].x = x;
                                        L5[n].width = width;
                                    }
                                    markHaveDisposedList.Add(key);
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static void 以某区域为参考对齐行和行高(CellJson cellJson, int row1, int row2, int sheet, List<CellInfo> L2)
        {
            List<string> markHaveDisposedList = new List<string>();
            for (var i = row1; i <= row2; i++)
            {
                var o = L2.Find(p => p.row == i);
                var y = o.y;
                var height = o.height;
                var L4 = cellJson.cells.FindAll(p => p.row == i && p.sheet == sheet);
                for (var k = 0; k < L4.Count; k++)
                {
                    var o1 = L4[k];
                    o1.y = y;
                    o1.height = height;
                    if (o1.cellPostionList != null)
                    {
                        var key1 = i + "_" + k + "_" + i;
                        if (!markHaveDisposedList.Contains(key1))
                        {
                            var L5 = o1.cellPostionList.FindAll(p => p.row == i);
                            for (var n = 0; n < L5.Count; n++)
                            {
                                L5[n].y = y;
                                L5[n].height = height;
                            }
                            markHaveDisposedList.Add(key1);
                        }
                    }
                    if (o1.cellPositionTotal != null)
                    {
                        o1.cellPositionTotal.y = y;
                    }
                    if (o1.mergeTo != null)
                    {
                        var c = o1.mergeTo.col;
                        var r = o1.mergeTo.row;
                        if (c != 0 && r != 0)
                        {
                            var oo = cellJson.cells.Find(p => p.col == c && p.row == r && p.sheet == sheet);
                            if (oo != null)
                            {
                                var key = c + "_" + r + "_" + i;
                                if (!markHaveDisposedList.Contains(key))
                                {
                                    var L5 = oo.cellPostionList.FindAll(p => p.row == i);
                                    for (var n = 0; n < L5.Count; n++)
                                    {
                                        L5[n].y = y;
                                        L5[n].height = height;
                                    }
                                    markHaveDisposedList.Add(key);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
