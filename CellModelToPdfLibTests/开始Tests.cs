using Microsoft.VisualStudio.TestTools.UnitTesting;
using CellModelToPdfLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;

namespace CellModelToPdfLib.Tests
{
    [TestClass()]
    public class 开始Tests
    {
        [TestMethod()]
        public void 开始处理Test()
        {
            string cellFile = Path.Combine(Environment.CurrentDirectory, "测试材料", "test.cll");
            using (CellToJson o = new CellToJson(cellFile))
            {
                var o1 = o.开始转换(0);
                string s = JsonConvert.SerializeObject(o1);
                File.WriteAllText(cellFile + ".json", s, Encoding.UTF8);
            }
        }
    }
}