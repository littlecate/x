using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.Model
{
    public class Formula
    {
        public int col { get; set; }
        public int row { get; set; }
        public int sheet { get; set; }
        public string str { get; set; }
    }
}
