using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class StringFormatX
    {
        public string fontFamily { get; set; }
        public double fontSize { get; set; }
        public int fontStyle { get; set; }
        public int fontColor { get; set; }
        public bool isMultiLine { get; set; }
        public bool isAutoScale { get; set; }
        public double lineSpace { get; set; }
    }
}
