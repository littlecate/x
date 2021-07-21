using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class CellInfo
    {
        public int col { get; set; }
        public int row { get; set; }
        public int sheet { get; set; }
        public int rowSpan { get; set; }       
        public bool isNewInsert { get; set; } = false;
        public int colSpan { get; set; }
        public double x { get; set; }
        public double y { get; set; }
        public double width { get; set; }
        public double height { get; set; }       
        public int cellAlign { get; set; }
        public string str { get; set; }
        public bool isHidden { get; set; }
        public CellPositionTotal cellPositionTotal { get; set; }
        public StringFormatX stringFormat { get; set; }
        public List<CellBorder> cellBorderList { get; set; }
        public List<CellPosition> cellPostionList { get; set; }
        public CellImage cellImage { get; set; }
        public bool isMergeCell { get; set; } = false;
        public CellInfo1 mergeTo { get; set; }
    }
}
