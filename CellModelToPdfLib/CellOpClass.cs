using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CellModelToPdfLib.CellOp;
using CellModelToPdfLib.Model;
using CSScriptLibrary;
using Newtonsoft.Json;

namespace CellModelToPdfLib
{
    public class CellOpClass
    {
        CellJson cellJson = null;
        CopyInfo copyInfo = null;
        public CellOpClass()
        {

        }

        public int OpenFile(string fileName, string password)
        {
            var t = File.ReadAllText(fileName, Encoding.UTF8);
            cellJson = JsonConvert.DeserializeObject<CellJson>(t);
            return 1;
        }

        public int SaveFile(string fileName, int closefile)
        {
            File.WriteAllText(fileName, JsonConvert.SerializeObject(cellJson), Encoding.UTF8);
            List<string> L = new List<string>();
            foreach (var p in cellJson.cells)
            {
                L.Add(JsonConvert.SerializeObject(p));
            }
            File.WriteAllText(fileName + ".txt", string.Join("\r\n\r\n", L));
            return 1;
        }

        public int closefile()
        {
            return 1;
        }

        public void SetCellString(int col, int row, int sheet, string str)
        {
            var t = S(col, row, sheet, str);
            if (t == 1)
            {
                CalculateSheet(sheet);
            }
        }

        public string GetCellString(int col, int row, int sheet)
        {
            var cell = GetCell(col, row, sheet);
            if (cell == null)
            {
                return "";
            }
            return cell.str;
        }

        public void InsertRow(int startRow, int count, int sheet)
        {
            var o = new InsertOneRowClass(cellJson);
            for (var i = 0; i < count; i++)
            {
                o.InsertOneRow(startRow, sheet);
            }
        }

        public void InsertCol(int startCol, int count, int sheet)
        {
            var o = new InsertOneColClass(cellJson);
            for (var i = 0; i < count; i++)
            {
                o.InsertOneCol(startCol, sheet);
            }
        }

        public void UnmergeCells(int col1, int row1, int col2, int row2, int cellBorderWidth = 2)
        {
            List<string> markList = new List<string>();
            var L = cellJson.cells.FindAll(p => p.col >= col1 && p.col <= col2 && p.row >= row1 && p.row <= row2
                                && p.isMergeCell == false && (p.rowSpan > 1 || p.colSpan > 1));
            for (var i = 0; i < L.Count; i++)
            {
                var p = L[i];
                var key = p.col + "_" + p.row;
                if (markList.Contains(key))
                {
                    continue;
                }
                markList.Add(key);
                foreach (var p1 in p.cellPostionList)
                {
                    var col = p1.col;
                    var row = p1.row;
                    var o = cellJson.cells.Find(p2 => p2.col == col && p2.row == row);
                    o.isMergeCell = false;
                    o.x = p1.x;
                    o.y = p1.y;
                    o.rowSpan = 1;
                    o.colSpan = 1;
                    o.mergeTo = new CellInfo1();
                    o.width = p1.width;
                    o.height = p1.height;
                    o.cellPositionTotal = new CellPositionTotal()
                    {
                        x = p1.x,
                        y = p1.y,
                        width = p1.width,
                        height = p1.height
                    };
                    o.cellPostionList = new List<CellPosition>()
                    {
                        new CellPosition()
                        {
                             x = p1.x,
                             y = p1.y,
                             width = p1.width,
                             height = p1.height,
                             col = p1.col,
                             row = p1.row
                        }
                    };
                    o.cellBorderList = new List<CellBorder>()
                    {
                        new CellBorder()
                        {
                            col = p1.col,
                            row = p1.row,
                            left = cellBorderWidth,
                            top = cellBorderWidth,
                            right = cellBorderWidth,
                            bottom = cellBorderWidth,
                            leftColor = 0,
                            topColor = 0,
                            rightColor = 0,
                            bottomColor = 0
                        }
                    };
                }
            }
        }

        public void MergeCells(int col1, int row1, int col2, int row2)
        {
            UnmergeCells(col1, row1, col2, row2, 2);
            List<string> markList = new List<string>();
            var L = cellJson.cells.FindAll(p => p.col >= col1 && p.col <= col2 && p.row >= row1 && p.row <= row2);
            var o = L[0];
            o.isMergeCell = false;
            o.colSpan = col2 - col1 + 1;
            o.rowSpan = row2 - row1 + 1;
            o.cellPositionTotal = new CellPositionTotal()
            {
                x = o.x,
                y = o.y,
                width = L.FindAll(p => p.row == o.row).Sum<CellInfo>(p => p.width),
                height = L.FindAll(p => p.col == o.col).Sum<CellInfo>(p => p.height),
            };
            var L1 = new List<CellPosition>();
            foreach (var p in L)
            {
                L1.AddRange(Comman.DeepCopyList<CellPosition>(p.cellPostionList));
            }
            o.cellPostionList = L1;
            CellBorder cellBorder1 = new CellBorder();
            if (o.row == row1)
            {
                cellBorder1.top = o.cellBorderList[0].top;
            }
            else if (o.row == row2)
            {
                cellBorder1.bottom = o.cellBorderList[0].bottom;
            }
            if (o.col == col1)
            {
                cellBorder1.left = o.cellBorderList[0].left;
            }
            if (o.col == col2)
            {
                cellBorder1.right = o.cellBorderList[0].right;
            }
            o.cellBorderList = new List<CellBorder>();
            o.cellBorderList.Add(cellBorder1);
            for (var i = 1; i < L.Count; i++)
            {
                var p = L[i];
                p.isMergeCell = true;
                p.cellPositionTotal = null;
                p.mergeTo = new CellInfo1()
                {
                    col = o.col,
                    row = o.row
                };
                CellBorder cellBorder = new CellBorder();
                if (p.row == row1)
                {
                    cellBorder.top = p.cellBorderList[0].top;
                }
                else if (p.row == row2)
                {
                    cellBorder.bottom = p.cellBorderList[0].bottom;
                }
                if (p.col == col1)
                {
                    cellBorder.left = p.cellBorderList[0].left;
                }
                if (p.col == col2)
                {
                    cellBorder.right = p.cellBorderList[0].right;
                }
                p.cellBorderList[0] = cellBorder;
                o.cellBorderList.AddRange(Comman.DeepCopyList<CellBorder>(p.cellBorderList));
            }
        }

        public void CopyRange(int col1, int row1, int col2, int row2)
        {
            var L = cellJson.cells.FindAll(p => p.col >= col1 && p.col <= col2 && p.row >= row1 && p.row <= row2);
            for (var i = 0; i < L.Count; i++)
            {
                L[i].isNewInsert = true;
            }
            copyInfo = new CopyInfo()
            {
                col1 = col1,
                row1 = row1,
                col2 = col2,
                row2 = row2,
                cells = Comman.DeepCopyList<CellInfo>(L)
            };
        }

        public void Paste(int col, int row, int sheet, int type, int sameSize, int skipBlank)
        {
            if (copyInfo == null)
            {
                return;
            }
            int chaCol = copyInfo.col1 - col;
            int chaRow = copyInfo.row1 - row;
            var col1 = col;
            var row1 = row;
            var col2 = copyInfo.col2 - chaCol;
            int row2 = copyInfo.row2 - chaRow;
            UnmergeCells(col1, row1, col2, row2);
            var L1 = Comman.DeepCopyList<CellInfo>(copyInfo.cells);
            var L3 = cellJson.cells.FindAll(p => p.col >= col1 && p.col <= col2 && p.row >= row1 && p.row <= row2);
            var L2 = Comman.DeepCopyList<CellInfo>(L3);
            double addX = L2[0].x - L1[0].x;
            double addY = L2[0].y - L1[0].y;
            for (var i = 0; i < L2.Count; i++)
            {
                L2[i] = Comman.CopyCellInfo(L1[i]);
                var o = L2[i];
                o.col -= chaCol;
                o.row -= chaRow;
                o.x += addX;
                o.y += addY;
                if (o.mergeTo != null && o.mergeTo.col != 0)
                {
                    o.mergeTo.col -= chaCol;
                    o.mergeTo.row -= chaRow;
                }
                o.cellPositionTotal.x += addX;
                o.cellPositionTotal.y += addY;
                for (var k = 0; k < o.cellPostionList.Count; k++)
                {
                    o.cellPostionList[k].x += addX;
                    o.cellPostionList[k].y += addY;
                    o.cellPostionList[k].col -= chaCol;
                    o.cellPostionList[k].row -= chaRow;
                }
                for (var k = 0; k < o.cellBorderList.Count; k++)
                {
                    o.cellBorderList[k].col -= chaCol;
                    o.cellBorderList[k].row -= chaRow;
                }
                var index = cellJson.cells.FindIndex(p => p.col == o.col && p.row == o.row && p.sheet == sheet);
                cellJson.cells.RemoveAt(index);
                cellJson.cells.Insert(index, Comman.CopyCellInfo(o));
            }
            for (var i = row1; i <= row2; i++)
            {
                var o = L2.Find(p => p.row == i);
                var y = o.y;
                var height = o.height;
                var L4 = cellJson.cells.FindAll(p => p.row == i && p.sheet == sheet);
                for (var k = 0; k < L4.Count; k++)
                {
                    var o1 = L4[k];
                    o1.y = y;
                    o1.height = height;
                    if (o1.cellPostionList != null)
                    {
                        var L5 = o1.cellPostionList.FindAll(p => p.row == i);
                        for (var n = 0; n < L5.Count; n++)
                        {
                            L5[n].y = y;
                            L5[n].height = height;
                        }
                    }
                }
            }
            for (var i = col1; i <= col2; i++)
            {
                var o = L2.Find(p => p.col == i);
                var x = o.x;
                var width = o.width;
                var L4 = cellJson.cells.FindAll(p => p.col == i && p.sheet == sheet);
                for (var k = 0; k < L4.Count; k++)
                {
                    var o1 = L4[k];
                    o1.x = x;
                    o1.width = width;
                    if (o1.cellPostionList != null)
                    {
                        var L5 = o1.cellPostionList.FindAll(p => p.col == i);
                        for (var n = 0; n < L5.Count; n++)
                        {
                            L5[n].x = x;
                            L5[n].width = width;
                        }
                    }
                }
            }
            for (var i = col1; i <= col2; i++)
            {
                for (var k = row1; k <= row2; k++)
                {
                    var o = cellJson.cells.Find(p => p.col == col1 && p.row == row1 && p.sheet == sheet);
                    if (o.cellPositionTotal != null)
                    {
                        o.cellPositionTotal.x = o.x;
                        o.cellPositionTotal.y = o.y;
                        o.cellPositionTotal.width = o.cellPostionList.FindAll(p => p.row == k).Sum<CellPosition>(p => p.width);
                        o.cellPositionTotal.height = o.cellPostionList.FindAll(p => p.col == i).Sum<CellPosition>(p => p.height);
                    }
                }
            }
        }

        private CellInfo GetCell(int col, int row, int sheet)
        {
            PageInfo page = GetColRowInWhichPage(col, row, sheet);
            if (page == null)
            {
                return null;
            }
            return cellJson.cells.Find(p => p.isMergeCell == false && p.col == col && p.row == row);
        }

        public void CalculateSheet(int sheet)
        {
            var L = cellJson.formulas.FindAll(p => p.sheet == sheet);
            var o = new CalFormulaClass(this);
            foreach (var p in L)
            {
                o.CalFormula(p, sheet);
            }
        }

        public int S(int col, int row, int sheet, string str)
        {
            var cell = GetCell(col, row, sheet);
            if (cell == null)
            {
                return -1;
            }
            cell.str = str;
            return 1;
        }

        private PageInfo GetColRowInWhichPage(int col, int row, int sheet)
        {
            return cellJson.pages.Find(p => p.sheet == sheet && p.startRow <= row && row <= p.endRow);
        }
    }
}
