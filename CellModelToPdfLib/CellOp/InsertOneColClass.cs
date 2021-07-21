using CellModelToPdfLib.Model;
using System;
using System.Collections.Generic;

namespace CellModelToPdfLib.CellOp
{
    internal class InsertOneColClass
    {
        private CellJson cellJson;

        public InsertOneColClass(CellJson cellJson)
        {
            this.cellJson = cellJson;
        }

        public void InsertOneCol(int startCol, int sheet)
        {
            var L = 拷贝前一列的数据(startCol, sheet);
            处理拷贝的数据列(L, startCol);
            插入新列(L, startCol);
            var addWidth = 得到插入的列的宽度(L[0], startCol);
            调整后面列的x值和列号(startCol, sheet, addWidth);
        }

        private void 处理拷贝的数据列(List<CellInfo> L, int startCol)
        {
            for (var i = 0; i < L.Count; i++)
            {
                var p = L[i];
                if (!p.isMergeCell)
                {
                    if (p.colSpan > 1)
                    {
                        p.isMergeCell = true;
                        p.width = p.cellPostionList[0].width;
                        p.mergeTo = new CellInfo1()
                        {
                            col = p.cellPostionList[0].col,
                            row = p.cellPostionList[0].row
                        };
                        p.cellPositionTotal.width = p.cellPostionList[0].width;
                        p.cellPostionList = Comman.DeepCopyList<CellPosition>(p.cellPostionList.FindAll(p1 => p1.col == startCol - 1));
                        p.cellBorderList = Comman.DeepCopyList<CellBorder>(p.cellBorderList.FindAll(p1 => p1.col == startCol - 1));
                    }
                    else
                    {
                        p.width = p.cellPostionList[0].width;
                        p.cellPositionTotal.width = p.cellPostionList[0].width;
                        p.cellPostionList = Comman.DeepCopyList<CellPosition>(p.cellPostionList.FindAll(p1 => p1.col == startCol - 1));
                        p.cellBorderList = Comman.DeepCopyList<CellBorder>(p.cellBorderList.FindAll(p1 => p1.col == startCol - 1));
                    }
                }
            }
        }

        private void 插入新列(List<CellInfo> L, int startCol)
        {
            bool isAdd = false;
            if (startCol == cellJson.cells.Count + 1)
            {
                isAdd = true;
            }
            for (var k = 0; k < L.Count; k++)
            {
                var p = L[k];
                p.isNewInsert = true;
                if (isAdd)
                {
                    cellJson.cells.Add(p);
                }
                else
                {
                    cellJson.cells.Insert(startCol - 1, p);
                }
            }
        }

        private void 调整后面列的x值和列号(int startCol, int sheet, double addWidth)
        {
            List<string> markList = new List<string>();
            var L2 = cellJson.cells.FindAll(p => p.col >= startCol && p.sheet == sheet);
            for (var i = 0; i < L2.Count; i++)
            {
                var p = L2[i];
                if (是否是合并单元格的第一个格子(p))
                {
                    var col = p.col;
                    var row = p.row;
                    var key = col + "_" + col;
                    markList.Add(key);
                    将格子位置整体右移一个宽度(p, addWidth);
                    将cellPositionList某列以右的格子右移(p, startCol - 1, addWidth);
                    调整cellBorderList列号(p, startCol - 1);
                }
                else
                {
                    var col = p.mergeTo.col;
                    var row = p.mergeTo.row;
                    if (col == 0 && col == 0)
                    {
                        continue;
                    }
                    var key = col + "_" + col;
                    if (markList.Contains(key))
                    {
                        continue; //如果格子已经处理过了，就不处理了
                    }
                    markList.Add(key);
                    var o = 得到合并到的格子(p, sheet);
                    调整mergeTo(p, startCol);
                    if (o.colSpan == 1)
                    {
                        调整CellPosition列号(o, startCol - 1);
                        调整cellBorderList列号(o, startCol - 1);
                        continue;
                    }
                    将格子整体加宽(o, addWidth);
                    增加cellPosition(o, startCol, addWidth);
                    增加cellBorder(o, startCol, addWidth);
                    将colSpan加1(o);
                }
            }
            调整单元格公式(startCol, sheet);
            将列号和x值右移(L2, addWidth);
            调整新插入列的属性(startCol, sheet, addWidth);
            cellJson.cols++;
        }

        private void 调整CellPosition列号(CellInfo p, int col)
        {
            var L4 = p.cellPostionList.FindAll(p1 => p1.col >= col);
            for (var n = 0; n < L4.Count; n++)
            {
                L4[n].col += 1;
            }
        }

        private void 调整单元格公式(int startCol, int sheet)
        {
            if (cellJson.formulas == null)
            {
                return;
            }
            var L = cellJson.formulas.FindAll(p => p.sheet == sheet);
            for (var i = 0; i < cellJson.formulas.Count; i++)
            {
                var p = cellJson.formulas[i];
                调整公式(p, startCol);
            }
        }

        private void 调整公式(Formula formula, int startCol)
        {
            List<string> L = Comman.拆分字串为词组(formula.str);
            List<string> L1 = new List<string>();
            List<string> L2 = new List<string>();
            foreach (var p in L)
            {
                if (Comman.IsColRowMark(p))
                {
                    int col = 0;
                    int row = 0;
                    Comman.GetColRowFromStrMark(p, ref col, ref row);
                    if (col < 65535 && row < 65536)
                    {
                        if (col > startCol)
                        {
                            col += 1;
                        }
                        var t = Comman.NumTo26(col) + "" + row;
                        L1.Add(p);
                        L2.Add(t);
                    }
                }
            }
            for (var m = 0; m < L1.Count; m++)
            {
                for (var i = 0; i < L.Count; i++)
                {
                    if (L[i] == L1[m])
                    {
                        L[i] = L2[m];
                    }
                }
            }
            if (formula.col > startCol)
            {
                formula.col += 1;
            }
            formula.str = string.Join(" ", L);
        }

        private void 调整mergeTo(CellInfo p, int startCol)
        {
            if (p.mergeTo.col >= startCol)
                p.mergeTo.col += 1;
        }

        private CellInfo 得到合并到的格子(CellInfo p, int sheet)
        {
            var col = p.mergeTo.col;
            var row = p.mergeTo.row;
            if (col == 0 && row == 0)
            {
                return null;
            }
            return cellJson.cells.Find(p1 => p1.col == col && p1.row == row && !p1.isNewInsert && p1.sheet == sheet);
        }

        private bool 是否是合并单元格的第一个格子(CellInfo p)
        {
            return !p.isMergeCell;
        }

        private void 将列号和x值右移(List<CellInfo> L2, double addWidth)
        {
            for (var i = 0; i < L2.Count; i++)
            {
                var p = L2[i];
                p.col += 1;
                p.x += addWidth;
            }
        }

        private void 调整新插入列的属性(int startCol, int sheet, double addWidth)
        {
            var L2 = cellJson.cells.FindAll(p => p.col == (startCol - 1) && p.isNewInsert && p.sheet == sheet);
            for (var i = 0; i < L2.Count; i++)
            {
                var p = L2[i];
                p.isNewInsert = false;
                p.col += 1;
                p.x += addWidth;
                if (p.cellPositionTotal != null)
                {
                    p.str = "xxxwww" + p.col;
                    p.cellPositionTotal.x += addWidth;
                    for (var n = 0; n < p.cellPostionList.Count; n++)
                    {
                        p.cellPostionList[n].x += addWidth;
                        p.cellPostionList[n].col += 1;
                    }
                    for (var n = 0; n < p.cellBorderList.Count; n++)
                    {
                        p.cellBorderList[n].col += 1;
                    }
                }
            }
        }

        private void 将colSpan加1(CellInfo o)
        {
            o.colSpan += 1;
        }

        private void 增加cellBorder(CellInfo o, int startCol, double addWidth)
        {
            将cellBorderList某列以右的格子右移(o, startCol, addWidth);
            var L5 = Comman.DeepCopyList<CellBorder>(o.cellBorderList.FindAll(p1 => p1.col == startCol - 1));
            var index = o.cellBorderList.FindIndex(p1 => p1.col == startCol + 1);
            for (var n = 0; n < L5.Count; n++)
            {
                L5[n].col += 1;
                L5[n].left = 0;
                L5[n].right = 0;
                o.cellBorderList.Insert(index, L5[n]);
            }
        }

        private void 将cellBorderList某列以右的格子右移(CellInfo o, int col, double addWidth)
        {
            var L3 = o.cellBorderList.FindAll(p1 => p1.col >= col);
            for (var n = 0; n < L3.Count; n++)
            {
                L3[n].col += 1;
            }
        }

        private void 增加cellPosition(CellInfo o, int startCol, double addWidth)
        {
            将cellPositionList某列以右的格子右移(o, startCol, addWidth);
            var L5 = Comman.DeepCopyList<CellPosition>(o.cellPostionList.FindAll(p1 => p1.col == startCol - 1));
            var index = o.cellPostionList.FindIndex(p1 => p1.col == startCol + 1);
            for (var n = 0; n < L5.Count; n++)
            {
                L5[n].col += 1;
                L5[n].x += addWidth;
                o.cellPostionList.Insert(index, L5[n]);
            }
        }

        private void 将格子整体加宽(CellInfo o, double addWidth)
        {
            o.cellPositionTotal.width += addWidth;
        }

        private void 调整cellBorderList列号(CellInfo p, int col)
        {
            var L4 = p.cellBorderList.FindAll(p1 => p1.col >= col);
            for (var n = 0; n < L4.Count; n++)
            {
                L4[n].col += 1;
            }
        }

        private void 将cellPositionList某列以右的格子右移(CellInfo p, int col, double addWidth)
        {
            var L3 = p.cellPostionList.FindAll(p1 => p1.col >= col);
            for (var n = 0; n < L3.Count; n++)
            {
                L3[n].x += addWidth;
                L3[n].col += 1;
            }
        }

        private void 将格子位置整体右移一个宽度(CellInfo p, double addWidth)
        {
            p.cellPositionTotal.x += addWidth;
        }

        private double 得到插入的列的宽度(CellInfo cellInfo, int startCol)
        {
            var width = 0d;
            if (cellInfo.cellPositionTotal != null)
            {
                width = cellInfo.cellPostionList.Find(p => p.col == startCol - 1).width;
            }
            else
            {
                width = cellInfo.width;
            }
            return width;
        }

        private List<CellInfo> 拷贝前一列的数据(int startCol, int sheet)
        {
            var L1 = cellJson.cells.FindAll(p => p.col == startCol - 1 && p.sheet == sheet);
            return Comman.DeepCopyList<CellInfo>(L1);
        }
    }
}