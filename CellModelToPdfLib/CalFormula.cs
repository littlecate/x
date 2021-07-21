using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib
{
    public class CalFormula
    {
        public string GetResult()
        {
            var o = new CalFormulaF();
            return ("具体内容占位符").ToString();
        }
    }


    public class CalFormulaF
    {       
        public string myif(bool condition, string s1, string s2)
        {
            if (condition)
            {
                return s1;
            }
            else
            {
                return s2;
            }
        }

        public double myvalue(string s)
        {
            try
            {
                return Convert.ToDouble(s);
            }
            catch
            {
                return 0;
            }
        }

        public int comp(string s1, string s2)
        {
            return string.Compare(s1, s2);
        }

        public int find(string s1, string s2)
        {
            return s1.IndexOf(s2);
        }

        public string left(string s1, int count)
        {
            if (count > s1.Length)
            {
                count = s1.Length;
            }
            return s1.Substring(0, count);
        }

        public string lower(string s)
        {
            return s.ToLower();
        }

        public string mid(string s, int start, int count)
        {
            if (start + count > s.Length)
            {
                count = s.Length - start;
            }
            return s.Substring(start, count);
        }

        public string right(string s, int count)
        {
            if (count > s.Length)
            {
                count = s.Length;
            }
            return s.Substring(s.Length - count);
        }

        public int strlen(string s)
        {
            return s.Length;
        }

        public string trimleft(string s)
        {
            return s.TrimStart();
        }

        public string trimright(string s)
        {
            return s.TrimEnd();
        }

        public string upper(string s)
        {
            return s.ToUpper();
        }

        public long date(int year, int month, int day)
        {
            return new DateTime(year, month, day).Ticks;
        }

        public string datestr(long date_serial, int style)
        {
            string dateFormat = "";
            DateTime date = new DateTime(date_serial);
            if (style == 0)
            {
                dateFormat = "yyyy-M-d";
            }
            else if (style == 1)
            {
                dateFormat = "yyyy-MM-dd HH:mm:ss";
            }
            else if (style == 2)
            {
                dateFormat = "yyyy-M-d H:m t\\M";
            }
            else if (style == 3)
            {
                dateFormat = "yyyy-MM-dd HH:mm";
            }
            else if (style == 4)
            {
                dateFormat = "yy-M-d";
            }
            else if (style == 5)
            {
                dateFormat = "M-d-yy";
            }
            else if (style == 6)
            {
                dateFormat = "MM-dd-yy";
            }
            else if (style == 7)
            {
                dateFormat = "M-d";
            }
            else if (style == 8)
            {
                return DateUpper.dateToUpper(date);
            }
            else if (style == 9)
            {
                var t = DateUpper.dateToUpper(date);
                int pos = t.IndexOf("月");
                return t.Substring(0, pos + 1);
            }
            else if (style == 10)
            {
                var t = DateUpper.dateToUpper(date);
                int pos = t.IndexOf("年");
                return t.Substring(pos + 1);
            }
            else if (style == 11)
            {
                dateFormat = "yyyy年M月d日";
            }
            else if (style == 12)
            {
                dateFormat = "yyyy年M月";
            }
            else if (style == 13)
            {
                dateFormat = "M月d日";
            }
            else if (style == 14)
            {
                return CaculateWeekDay(date);
            }
            else if (style == 15)
            {
                return CaculateWeekDay(date).Replace("星期", "");
            }
            else if (style == 16)
            {
                dateFormat = "d-MMM";
            }
            else if (style == 17)
            {
                dateFormat = "d-MMM-yy";
            }
            else if (style == 18)
            {
                dateFormat = "dd-MMM-yy";
            }
            else if (style == 19)
            {
                dateFormat = "MMM-yy";
            }
            else if (style == 20)
            {
                dateFormat = "MMMM-yy";
            }
            return date.ToString(dateFormat);
        }

        public long datevalue(string s)
        {
            return Convert.ToDateTime(s).Ticks;
        }

        public int day(string s)
        {
            var t = ToDate(s);
            return t.Day;
        }

        public int hour(string s)
        {
            var t = ToDate(s);
            return t.Hour;
        }

        public int minute(string s)
        {
            var t = ToDate(s);
            return t.Minute;
        }

        public int month(string s)
        {
            var t = ToDate(s);
            return t.Month;
        }

        public long now()
        {
            return DateTime.Now.Ticks;
        }

        public int second(string s)
        {
            var t = ToDate(s);
            return t.Second;
        }

        public long today()
        {
            return now();
        }

        public int weekday(string s, int type)
        {
            var t = ToDate(s);
            return (int)t.DayOfWeek;
        }

        public int year(string s)
        {
            var t = ToDate(s);
            return t.Year;
        }

        public int yearday(string s)
        {
            var t = ToDate(s);
            return t.DayOfYear;
        }

        public string mystring(double v)
        {
            return v.ToString();
        }

        public string mystring(string v)
        {
            return v.ToString();
        }

        public string cnnum(string s, string style)
        {
            throw new NotImplementedException();
        }

        private DateTime ToDate(string s)
        {
            if (Comman.IsNumber(s))
            {
                return new DateTime(Convert.ToInt32(s));
            }
            else
            {
                return Convert.ToDateTime(s);
            }
        }

        /// <summary>
        /// 计算日期星期几
        /// </summary>
        /// <param name="date">日期（DateTime类型）</param>
        /// <returns></returns>
        public string CaculateWeekDay(DateTime date)
        {
            var weekdays = new string[] { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            return weekdays[(int)date.DayOfWeek];
        }
    }
}
