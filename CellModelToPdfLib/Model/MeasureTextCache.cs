using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class MeasureTextCache
    {
        public string text { get; set; }
        public XFont font { get; set; }
        public XSize size { get; set; }

    }
}
