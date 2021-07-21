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
    public class JsonToPdfTests
    {
        [TestMethod()]
        public void startTest()
        {
            string cellFile = Path.Combine(Environment.CurrentDirectory, "测试材料", "test.cll");
            using (CellToJson o1 = new CellToJson(cellFile))
            {
                var o2 = o1.开始转换(0);
                string s1 = JsonConvert.SerializeObject(o2);
                File.WriteAllText(cellFile + ".json", s1, Encoding.UTF8);
            }
            string jsonFile = Path.Combine(Environment.CurrentDirectory, "测试材料", "test.cll.json");
            string s = File.ReadAllText(jsonFile, Encoding.UTF8);
            var cellJson = JsonConvert.DeserializeObject<Model.CellJson>(s);
            var o = new JsonToPdf(cellJson, @"G:\项目\CellModelToPdf\CellModelToPdfLib\bin\Debug");
            o.start();
        }
    }
}