using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class CopyInfo
    {
        public int col1 { get; set; }
        public int row1 { get; set; }
        public int col2 { get; set; }
        public int row2 { get; set; }
        public List<CellInfo> cells { get; set; }
    }
}
