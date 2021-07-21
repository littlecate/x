using CellModelToPdfLib.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib.CellOp
{
    public class InsertOneRowClass
    {
        CellJson cellJson;
        public InsertOneRowClass(CellJson _cellJson)
        {
            cellJson = _cellJson;
        }

        public void InsertOneRow(int startRow, int sheet)
        {
            var L = 拷贝前一行的数据(startRow, sheet);
            处理拷贝的数据行(L, startRow);
            插入新行(L, startRow);
            var addHeight = 得到插入的行的高度(L[0], startRow);
            调整后面行的y值和行号(startRow, sheet, addHeight);
        }

        private void 处理拷贝的数据行(List<CellInfo> L, int startRow)
        {
            for (var i = 0; i < L.Count; i++)
            {
                var p = L[i];
                if (!p.isMergeCell)
                {
                    if (p.rowSpan > 1)
                    {
                        p.isMergeCell = true;
                        p.height = p.cellPostionList[0].height;
                        p.mergeTo = new CellInfo1()
                        {
                            col = p.cellPostionList[0].col,
                            row = p.cellPostionList[0].row
                        };
                        p.cellPositionTotal.height = p.cellPostionList[0].height;
                        p.cellPostionList = Comman.DeepCopyList<CellPosition>(p.cellPostionList.FindAll(p1 => p1.row == startRow - 1));
                        p.cellBorderList = Comman.DeepCopyList<CellBorder>(p.cellBorderList.FindAll(p1 => p1.row == startRow - 1));
                    }
                    else
                    {
                        p.height = p.cellPostionList[0].height;
                        p.cellPositionTotal.height = p.cellPostionList[0].height;
                        p.cellPostionList = Comman.DeepCopyList<CellPosition>(p.cellPostionList.FindAll(p1 => p1.row == startRow - 1));
                        p.cellBorderList = Comman.DeepCopyList<CellBorder>(p.cellBorderList.FindAll(p1 => p1.row == startRow - 1));
                    }
                }
            }
        }

        private void 插入新行(List<CellInfo> L, int startRow)
        {
            bool isAdd = false;
            if (startRow == cellJson.cells.Count + 1)
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
                    cellJson.cells.Insert(startRow - 1, p);
                }
            }
        }

        private void 调整后面行的y值和行号(int startRow, int sheet, double addHeight)
        {
            List<string> markList = new List<string>();
            var L2 = cellJson.cells.FindAll(p => p.row >= startRow && p.sheet == sheet);
            for (var i = 0; i < L2.Count; i++)
            {
                var p = L2[i];
                if (是否是合并单元格的第一个格子(p))
                {
                    var col = p.col;
                    var row = p.row;
                    var key = col + "_" + row;
                    markList.Add(key);
                    将格子位置整体下移一个高度(p, addHeight);
                    将cellPositionList某行以下的格子下移(p, startRow - 1, addHeight);
                    调整cellBorderList行号(p, startRow - 1);
                }
                else
                {
                    var col = p.mergeTo.col;
                    var row = p.mergeTo.row;
                    if (col == 0 && row == 0)
                    {
                        continue;
                    }
                    var key = col + "_" + row;
                    if (markList.Contains(key))
                    {
                        continue; //如果格子已经处理过了，就不处理了
                    }
                    markList.Add(key);
                    var o = 得到合并到的格子(p, sheet);
                    调整mergeTo(p, startRow);
                    if (o.rowSpan == 1)
                    {
                        调整CellPosition行号(o, startRow - 1);
                        调整cellBorderList行号(o, startRow - 1);
                        continue;
                    }
                    将格子整体加高(o, addHeight);
                    增加cellPosition(o, startRow, addHeight);
                    增加cellBorder(o, startRow, addHeight);
                    将rowSpan加1(o);
                }
            }
            调整单元格公式(startRow, sheet);
            将行号和y值下移(L2, addHeight);
            调整新插入行的属性(startRow, sheet, addHeight);
            cellJson.rows++;
        }

        private void 调整CellPosition行号(CellInfo p, int row)
        {
            var L4 = p.cellPostionList.FindAll(p1 => p1.row >= row);
            for (var n = 0; n < L4.Count; n++)
            {
                L4[n].row += 1;
            }
        }

        private void 调整单元格公式(int startRow, int sheet)
        {
            if (cellJson.formulas == null)
            {
                return;
            }
            var L = cellJson.formulas.FindAll(p => p.sheet == sheet);
            for (var i = 0; i < cellJson.formulas.Count; i++)
            {
                var p = cellJson.formulas[i];
                调整公式(p, startRow);
            }
        }

        private void 调整公式(Formula formula, int startRow)
        {
            List<string> L = Comman.拆分字串为词组(formula.str);
            List<string> L1 = new List<string>();
            List<string> L2 = new List<string>();
            foreach (var p in L)
            {
                if (Comman.IsColRowMark(p))
                {
                    string col = "";
                    int row = 0;
                    Comman.GetColRowFromStrMark(p, ref col, ref row);
                    if (row < 65535 && col.Length <= 3)
                    {
                        if (row > startRow)
                        {
                            row += 1;
                        }
                        var t = col + "" + row;
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
            if (formula.row > startRow)
            {
                formula.row += 1;
            }
            formula.str = string.Join(" ", L);
        }

        private void 调整mergeTo(CellInfo p, int startRow)
        {
            if (p.mergeTo.row >= startRow)
                p.mergeTo.row += 1;
        }

        private CellInfo 得到合并到的格子(CellInfo p, int sheet)
        {
            var col = p.mergeTo.col;
            var row = p.mergeTo.row;
            if (col == 0 && row == 0)
            {
                return null;
            }
            return cellJson.cells.Find(p1 => p1.row == row && p1.col == col && !p1.isNewInsert && p1.sheet == sheet);
        }

        private bool 是否是合并单元格的第一个格子(CellInfo p)
        {
            return !p.isMergeCell;
        }

        private void 将行号和y值下移(List<CellInfo> L2, double addHeight)
        {
            for (var i = 0; i < L2.Count; i++)
            {
                var p = L2[i];
                p.row += 1;
                p.y += addHeight;
            }
        }

        private void 调整新插入行的属性(int startRow, int sheet, double addHeight)
        {
            var L2 = cellJson.cells.FindAll(p => p.row == (startRow - 1) && p.isNewInsert && p.sheet == sheet);
            for (var i = 0; i < L2.Count; i++)
            {
                var p = L2[i];
                p.isNewInsert = false;
                p.row += 1;
                p.y += addHeight;
                if (p.cellPositionTotal != null)
                {
                    p.str = "xxxwww" + p.row;
                    p.cellPositionTotal.y += addHeight;
                    for (var n = 0; n < p.cellPostionList.Count; n++)
                    {
                        p.cellPostionList[n].y += addHeight;
                        p.cellPostionList[n].row += 1;
                    }
                    for (var n = 0; n < p.cellBorderList.Count; n++)
                    {
                        p.cellBorderList[n].row += 1;
                    }
                }
            }
        }

        private void 将rowSpan加1(CellInfo o)
        {
            o.rowSpan += 1;
        }

        private void 增加cellBorder(CellInfo o, int startRow, double addHeight)
        {
            将cellBorderList某行以下的格子下移(o, startRow, addHeight);
            var L5 = Comman.DeepCopyList<CellBorder>(o.cellBorderList.FindAll(p1 => p1.row == startRow - 1));
            var index = o.cellBorderList.FindIndex(p1 => p1.row == startRow + 1);
            for (var n = 0; n < L5.Count; n++)
            {
                L5[n].row += 1;
                L5[n].top = 0;
                L5[n].bottom = 0;
                o.cellBorderList.Insert(index, L5[n]);
            }
        }

        private void 将cellBorderList某行以下的格子下移(CellInfo o, int row, double addHeight)
        {
            var L3 = o.cellBorderList.FindAll(p1 => p1.row >= row);
            for (var n = 0; n < L3.Count; n++)
            {
                L3[n].row += 1;
            }
        }

        private void 增加cellPosition(CellInfo o, int startRow, double addHeight)
        {
            将cellPositionList某行以下的格子下移(o, startRow, addHeight);
            var L5 = Comman.DeepCopyList<CellPosition>(o.cellPostionList.FindAll(p1 => p1.row == startRow - 1));
            var index = o.cellPostionList.FindIndex(p1 => p1.row == startRow + 1);
            for (var n = 0; n < L5.Count; n++)
            {
                L5[n].row += 1;
                L5[n].y += addHeight;
                o.cellPostionList.Insert(index, L5[n]);
            }
        }

        private void 将格子整体加高(CellInfo o, double addHeight)
        {
            o.cellPositionTotal.height += addHeight;
        }

        private void 调整cellBorderList行号(CellInfo p, int row)
        {
            var L4 = p.cellBorderList.FindAll(p1 => p1.row >= row);
            for (var n = 0; n < L4.Count; n++)
            {
                L4[n].row += 1;
            }
        }

        private void 将cellPositionList某行以下的格子下移(CellInfo p, int row, double addHeight)
        {
            var L3 = p.cellPostionList.FindAll(p1 => p1.row >= row);
            for (var n = 0; n < L3.Count; n++)
            {
                L3[n].y += addHeight;
                L3[n].row += 1;
            }
        }

        private void 将格子位置整体下移一个高度(CellInfo p, double addHeight)
        {
            p.cellPositionTotal.y += addHeight;
        }

        private double 得到插入的行的高度(CellInfo cellInfo, int startRow)
        {
            var height = 0d;
            if (cellInfo.cellPositionTotal != null)
            {
                height = cellInfo.cellPostionList.Find(p => p.row == startRow - 1).height;
            }
            else
            {
                height = cellInfo.height;
            }
            return height;
        }

        private List<CellInfo> 拷贝前一行的数据(int startRow, int sheet)
        {
            var L1 = cellJson.cells.FindAll(p => p.row == startRow - 1 && p.sheet == sheet);
            return Comman.DeepCopyList<CellInfo>(L1);
        }

    }
}
