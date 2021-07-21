using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib
{
    public class GlobalV
    {
        /// <summary>
        /// 细线宽
        /// </summary>
        public static double line1w = 0;

        /// <summary>
        /// 中线宽
        /// </summary>
        public static double line2w = 0;

        /// <summary>
        /// 粗线宽
        /// </summary>
        public static double line3w = 0;

        /// <summary>
        /// 单元格默认高度
        /// </summary>
        public static double cellDefaultHeight = 20;

        public static List<char> canNotInTheFirstChars = new List<char>()
        {
            ',',
            '.',
            '，',
            '。',
            '；',
            ';',
            '！',
            ',',
            '‰',
            ',',
            '!',
            '：',
            '〗',
            '％',
            //'＄', 
			'＇',
            '｝',
            '］',
            '》',
            '】',
            '』',
            '＠',
            //'＝',
            //'≥',            
			//'≤',
            //'＞',
            '、',
            '）',
            '?' ,
            '？',
            '．',
            '）',
            ')',
            '»',
            '›',
            //'〉',
            '﹞',
            //'>',
            '”',
            '’',
            //'〞',
        };

        public static List<char> canNotInTheEndChars = new List<char>()
        {
            '＜',
            '{',
            '[' ,
            '（',
            '〔',
            '【',
            '《',
            '「',
            '『',
            '〖',
            //'"',
            //'＂',
            '｛',
            '［',
            '<',
            '«',
            '‹',
            '﹝',
            '〈',
            '“',
            '‘',
            '〝',
        };

        public static List<char> 不能分开的字符 = new List<char>()
        {
            '0','1','2','3','4','5','6','7','8','9','/','%','‰',
            'a','b','c','d','e','f','g',
            'h','i','j','k','l','m','n',
            'o','p','q',
            'r','s','t',
            'u','v','w',
            'x','y','z',
            'A','B','C','D','E','F','G',
            'H','I','J','K','L','M','N',
            'O','P','Q',
            'R','S','T',
            'U','V','W',
            'X','Y','Z'
        };

        public static List<char> 数字 = new List<char>()
        {
            '0','1','2','3','4','5','6','7','8','9','.'
        };

        public static List<char> 标准号不可拆分字符 = new List<char>()
        {
            '0','1','2','3','4','5','6','7','8','9','.','G','B','T','/','-',' ','—','　'
        };
    }
}
