using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class CellJson
    {
        public List<PageInfo> pages { get; set; }
        public List<ImageInfo> images { get; set; }
        public BackGroundImageInfo backGroundImageInfo { get; set; }
        public List<Formula> formulas { get; set; }       
        public List<FloatImage> floatImages { get; set; }
        public List<CellInfo> cells { get; set; }
        public int cols { get; set; }
        public int rows { get; set; }
    }
}
