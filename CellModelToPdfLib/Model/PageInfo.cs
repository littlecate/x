using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class PageInfo
    {
        public int sheet { get; set; }
        public double paperWidth { get; set; }
        public double paperHeight { get; set; }
        public double contentWidth { get; set; }
        public double contentHeight { get; set; }
        public int startRow { get; set; }
        public int endRow { get; set; }
        public int marginLeft { get; set; }
        public int marginTop { get; set; }
        public int marginRight { get; set; }
        public int marginBottom { get; set; }              
    }
}
