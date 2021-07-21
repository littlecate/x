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
    public class CellOpClassTests
    {
        [TestMethod()]
        public void SetCellStringTest()
        {
            string cellFile = Path.Combine(Environment.CurrentDirectory, "测试材料", "test.cll");
            using (CellToJson o = new CellToJson(cellFile))
            {
                var o1 = o.开始转换(0);
                string s = JsonConvert.SerializeObject(o1);
                File.WriteAllText(cellFile + ".json", s, Encoding.UTF8);
            }
            string jsonFile = Path.Combine(Environment.CurrentDirectory, "测试材料", "test.cll.json");
            string jsonFile1 = Path.Combine(Environment.CurrentDirectory, "测试材料", "test.cll.json.json");
            CellOpClass cellOpClass = new CellOpClass();
            cellOpClass.OpenFile(jsonFile, "");
            //cellOpClass.SetCellString(6, 10, 0, "测试字符串");
            //cellOpClass.InsertRow(8, 2, 0);
            //cellOpClass.SetCellString(2, 9 + 2, 0, "testbbbb");
            //cellOpClass.InsertCol(2, 2, 0);
            //cellOpClass.UnmergeCells(1, 1, 22, 2);
            //cellOpClass.MergeCells(6, 5, 17, 9);
            cellOpClass.CopyRange(6, 5, 17, 9);
            cellOpClass.Paste(6, 12, 0, 0, 1, 0);
            cellOpClass.SaveFile(jsonFile1, 1);
            cellOpClass.closefile();
            ConvertToPdf(jsonFile1);
        }

        private void ConvertToPdf(string jsonFile)
        {
            string s = File.ReadAllText(jsonFile, Encoding.UTF8);
            var cellJson = JsonConvert.DeserializeObject<Model.CellJson>(s);
            var o = new JsonToPdf(cellJson, @"D:\项目\CellModelToPdf\CellModelToPdfLib\bin\Debug");
            o.start(@"D:\项目\CellModelToPdf\CellModelToPdfLibTests\bin\Debug\测试材料\test.pdf");
        }
    }
}