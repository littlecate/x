using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class CellBorder
    {
        public int col { get; set; }
        public int row { get; set; }
        public int left { get; set; } = 0;
        public int top { get; set; } = 0;
        public int right { get; set; } = 0;
        public int bottom { get; set; } = 0;
        public int leftColor { get; set; }
        public int topColor { get; set; }
        public int rightColor { get; set; }
        public int bottomColor { get; set; }
    }
}
