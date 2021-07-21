using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class XFontDrawCache
    {
        public string fontFamily { get; set; }
        public double fontSize { get; set; }
        public XFontStyle xFontStyle { get; set; }
        public XFont xFont { get; set; }
    }
}
