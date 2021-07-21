using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CELL50LibU;
using CellModelToPdfLib.Model;
using Newtonsoft.Json;

namespace CellModelToPdfLib
{
    public class CellToJson : IDisposable
    {
        CELL50LibU.CellClass cell = null;
        string fileName;
        int unitType = 0;
        List<MyPoint> haveDisposeedCellList = new List<MyPoint>();
        int marginLeft, marginTop, marginRight, marginBottom, printHAlign, printVAlign;
        double paperWidth, paperHeight;
        int contentWidth, contentHeight;
        List<ImageInfo> imageInfoList = new List<ImageInfo>();
        List<Formula> formulas = new List<Formula>();
        List<CellInfo> cells = new List<CellInfo>();
        List<string> tempFileList = new List<string>();
        string imageTempFileName = "";
        int type = 1;
        /// <summary>
        /// x起始页面偏移
        /// </summary>
        double xPy = 0;
        /// <summary>
        /// y起始页面偏移
        /// </summary>
        double yPy = 0;
        public CellToJson(string fileName)
        {
            type = 1;
            this.fileName = fileName;
            imageTempFileName = Guid.NewGuid().ToString();
            cell = new CELL50LibU.CellClass();
            cell.OpenFile(fileName, "");

            marginLeft = (int)mmmToPixel(cell.PrintGetMargin(0));
            marginTop = (int)mmmToPixel(cell.PrintGetMargin(1));
            marginRight = (int)mmmToPixel(cell.PrintGetMargin(2));
            marginBottom = (int)mmmToPixel(cell.PrintGetMargin(3));
        }

        public CellToJson(CellClass _cell)
        {
            type = 2;
            imageTempFileName = Guid.NewGuid().ToString();
            cell = _cell;
            marginLeft = (int)mmmToPixel(cell.PrintGetMargin(0));
            marginTop = (int)mmmToPixel(cell.PrintGetMargin(1));
            marginRight = (int)mmmToPixel(cell.PrintGetMargin(2));
            marginBottom = (int)mmmToPixel(cell.PrintGetMargin(3));
        }

        public void Dispose()
        {
            if (type == 1)
            {
                cell.closefile();
            }
            foreach (var p in tempFileList)
            {
                File.Delete(p); //删除不掉，郁闷中....
            }
        }

        public CellJson 开始转换(int sheet)
        {
            printHAlign = cell.PrintGetHAlign(sheet);
            printVAlign = cell.PrintGetVAlign(sheet);

            paperWidth = mmmToPixel(cell.PrintGetPaperWidth(sheet));
            paperHeight = mmmToPixel(cell.PrintGetPaperHeight(sheet));

            List<int> pageBreakRowList = GetPageBreakRowList();
            List<PageInfo> pageInfos = GetPageInfos(pageBreakRowList);

            GetPageInfoList(pageInfos, sheet);

            CellJson o = new CellJson()
            {
                pages = pageInfos,
                cells = cells,
                backGroundImageInfo = GetBackGroundImageInfo(sheet),
                floatImages = GetFloatImages(sheet),
                images = imageInfoList,
                formulas = formulas,
                cols = cell.GetCols(sheet),
                rows = cell.GetRows(sheet)
            };
            return o;
        }

        private int GetContentHeight(int startRow, int endRow, int sheet)
        {
            var h = 0;
            for (var i = startRow; i <= endRow; i++)
            {
                h += cell.GetRowHeight(unitType, i, sheet);
            }
            return h;
        }

        private int GetContentWidth(int startCol, int endCol, int sheet)
        {
            var w = 0;
            for (var i = startCol; i <= endCol; i++)
            {
                w += cell.GetColWidth(unitType, i, sheet);
            }
            return w;
        }

        private void GetPageInfoList(List<PageInfo> pageInfos, int sheet)
        {
            foreach (var pageInfo in pageInfos)
            {
                contentWidth = GetContentWidth(1, cell.GetCols(sheet) - 1, sheet);
                contentHeight = GetContentHeight(pageInfo.startRow, pageInfo.endRow, sheet);
                xPy = GetXPy();
                yPy = GetYPy();
                for (var row = pageInfo.startRow; row <= pageInfo.endRow; row++)
                {
                    for (var col = 1; col < cell.GetCols(sheet); col++)
                    {
                        var o1 = haveDisposeedCellList.Find(p => p.X == col && p.Y == row);
                        if (o1 != null && o1.X != 0 && o1.Y != 0)
                        {
                            continue;
                        }
                        string formula = cell.GetFormula(col, row, sheet);
                        if (!string.IsNullOrEmpty(formula))
                        {
                            formulas.Add(new Formula()
                            {
                                col = col,
                                row = row,
                                sheet = sheet,
                                str = formula
                            });
                        }
                        List<Model.CellBorder> cellBorderList = new List<Model.CellBorder>();
                        List<Model.CellPosition> cellPositionList = new List<Model.CellPosition>();
                        var x = GetCellX(1, col, sheet);
                        var y = GetCellY(pageInfo.startRow, row, sheet);
                        int width = 0;
                        int height = 0;
                        List<CellInfo> L = new List<CellInfo>();
                        GetCellProp(col, row, sheet, ref cellBorderList, ref width, ref height, ref cellPositionList, pageInfo, ref L);
                        L[0].col = col;
                        L[0].row = row;
                        L[0].sheet = sheet;                       
                        L[0].cellAlign = cell.GetCellAlign(col, row, sheet);
                        L[0].str = cell.GetCellString(col, row, sheet).TrimEnd(new char[] { '\r', '\n' }); //尾部的回车不起作用，故去除
                        L[0].stringFormat = new Model.StringFormatX()
                        {
                            fontFamily = cell.GetFontName(cell.GetCellFont(col, row, sheet)),
                            fontSize = cell.GetCellFontSize(col, row, sheet),
                            fontStyle = cell.GetCellFontStyle(col, row, sheet),
                            fontColor = cell.GetColor(cell.GetCellTextColor(col, row, sheet)),
                            lineSpace = mmmToPixels2(cell.GetCellTextLineSpace(col, row, sheet) * 10)
                        };
                        L[0].isHidden = (width == 0 || height == 0);
                        L[0].cellPositionTotal = new Model.CellPositionTotal()
                        {                           
                            x = mmmToPixels2(x),
                            y = mmmToPixels2(y),
                            width = mmmToPixels2(width),
                            height = mmmToPixels2(height)
                        };
                        L[0].cellBorderList = cellBorderList;
                        L[0].cellPostionList = cellPositionList;
                        L[0].cellImage = GetCellImageInfo(col, row, sheet);
                        cells.AddRange(L);
                    }
                }
                pageInfo.sheet = sheet;
                pageInfo.paperWidth = mmmToPixels2(pageInfo.paperWidth);
                pageInfo.paperHeight = mmmToPixels2(pageInfo.paperHeight);
                pageInfo.contentWidth = mmmToPixels2(contentWidth);
                pageInfo.contentHeight = mmmToPixels2(contentHeight);
                pageInfo.marginLeft = (int)mmmToPixels2(marginLeft);
                pageInfo.marginRight = (int)mmmToPixels2(marginRight);
                pageInfo.marginTop = (int)mmmToPixels2(marginTop);
                pageInfo.marginBottom = (int)mmmToPixels2(marginBottom);
            }
        }

        private List<FloatImage> GetFloatImages(int sheet)
        {
            List<FloatImage> L = new List<FloatImage>();
            string name = cell.GetFirstFloatImage(sheet);
            double rate = 72f / 96f;
            while (!string.IsNullOrEmpty(name))
            {
                int xpos = 0, ypos = 0, width = 0, height = 0;
                var t = cell.GetFloatImagePos(sheet, name, ref xpos, ref ypos, ref width, ref height);
                if (t > 0)
                {
                    int index = cell.GetFloatImageAttribute(sheet, name, 5);
                    将图片添加到imageInfoList(index);
                    L.Add(new FloatImage()
                    {
                        index = index,
                        name = name,
                        xpos = xpos * rate,
                        ypos = ypos * rate,
                        width = width * rate,
                        height = height * rate
                    });
                }
                name = cell.GetNextFloatImage(sheet);
            }
            return L;
        }

        private BackGroundImageInfo GetBackGroundImageInfo(int sheet)
        {
            int index = 0, style = 0;
            int t = cell.GetBackImage(sheet, ref index, ref style);
            if (t == 1)
            {
                将图片添加到imageInfoList(index);
                return new BackGroundImageInfo()
                {
                    index = index,
                    style = style
                };
            }
            return null;
        }

        private CellImage GetCellImageInfo(int col, int row, int sheet)
        {
            int index = 0, style = 0, valign = 0, halign = 0;
            int t = cell.GetCellImage(col, row, sheet, ref index, ref style, ref halign, ref valign);
            if (t == 1)
            {
                将图片添加到imageInfoList(index);
                return new CellImage()
                {
                    index = index,
                    style = style,
                    halign = halign,
                    valign = valign
                };
            }
            return null;
        }

        private void 将图片添加到imageInfoList(int index)
        {
            if (imageInfoList.Find(p => p.index == index) == null)
            {
                List<string> 文件后缀列表 = new List<string>() { "jpg", "bmp", "gif" };
                foreach (var p in 文件后缀列表)
                {
                    string f = Path.Combine(Environment.CurrentDirectory, "temp", imageTempFileName + "." + p);
                    string path = Path.GetDirectoryName(f);
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    if (cell.SaveImageFile(index, f))
                    {
                        imageInfoList.Add(new ImageInfo()
                        {
                            index = index,
                            image = File.ReadAllBytes(f)
                        });
                        break;
                    }
                    tempFileList.Add(f);
                }
            }
        }

        private List<PageInfo> GetPageInfos(List<int> pageBreakRowList)
        {
            List<PageInfo> L = new List<PageInfo>();
            for (var i = 0; i < pageBreakRowList.Count; i++)
            {
                int startRow = 0;
                int endRow = 0;
                if (i == 0)
                {
                    startRow = 1;
                    endRow = pageBreakRowList[i];
                }
                else
                {
                    startRow = pageBreakRowList[i - 1] + 1;
                    endRow = pageBreakRowList[i];
                }
                L.Add(new PageInfo()
                {
                    paperWidth = paperWidth,
                    paperHeight = paperHeight,
                    startRow = startRow,
                    endRow = endRow
                });
            }
            return L;
        }

        private List<int> GetPageBreakRowList()
        {
            var h = 0;
            var pageHeight = GetPageHeight();
            List<int> L = new List<int>();
            for (var row = 1; row < cell.GetRows(0); row++)
            {
                h += cell.GetRowHeight(0, row, 0);
                if (h > pageHeight || cell.IsRowPageBreak(row) == 1)
                {
                    if (row == 1)
                    {
                        //
                    }
                    else
                    {
                        L.Add(row - 1);
                    }
                    h = cell.GetRowHeight(0, row, 0);
                }
            }
            if (L.Count == 0 || L[L.Count - 1] != cell.GetRows(0))
            {
                L.Add(cell.GetRows(0) - 1);
            }
            return L;
        }

        private int GetPageHeight()
        {
            int paperHeight = cell.PrintGetPaperHeight(0);
            int marginTop = cell.PrintGetMargin(1);
            int marginBottom = cell.PrintGetMargin(3);
            int pageHeight = paperHeight - marginTop - marginBottom;
            return pageHeight;
        }

        /// <summary>
        /// 得到x起始页面偏移
        /// </summary>
        /// <returns></returns>
        private double GetXPy()
        {
            if (printHAlign == 1) //水平居中
            {
                return (paperWidth - marginLeft - marginRight - contentWidth) / 2 + marginLeft;
            }
            return marginLeft;
        }

        private double GetYPy()
        {
            if (printVAlign == 1) //垂直居中
            {
                return (paperHeight - marginTop - marginBottom - contentHeight) / 2 + marginTop;
            }
            return marginTop;
        }

        private double mmmToPixel(int v)
        {
            if (unitType == 0) //不作转换
            {
                return v;
            }
            return (96 * (1 / 25.4f)) * (v / 10);
        }

        private void GetCellProp(int col, int row, int sheet, ref List<Model.CellBorder> cellBorderList, ref int width, ref int height,
            ref List<Model.CellPosition> cellPositionList, PageInfo pageInfo, ref List<CellInfo> L)
        {
            cellBorderList.Clear();
            cellPositionList.Clear();
            CellInfo t = new CellInfo();
            int c1 = 0, r1 = 0, c2 = 0, r2 = 0;
            if (cell.GetMergeRange(col, row, ref c1, ref r1, ref c2, ref r2) == 1)
            {
                var w = 0;
                for (var i = c1; i <= c2; i++)
                {
                    w += cell.GetColWidth(unitType, i, sheet);
                }
                width = w;
                var h = 0;
                for (var i = r1; i <= r2; i++)
                {
                    h += cell.GetRowHeight(unitType, i, sheet);
                }
                height = h;
                for (var i = r1; i <= r2; i++)
                {
                    for (var k = c1; k <= c2; k++)
                    {
                        haveDisposeedCellList.Add(new MyPoint()
                        {
                            X = k,
                            Y = i
                        });
                        CellInfo t1 = new CellInfo();
                        if (k == c1 && i == r1)
                        {
                            t1.isMergeCell = false;
                            t1.col = k;
                            t1.row = i;                           
                            t1.rowSpan = r2 - r1 + 1;
                            t1.colSpan = c2 - c1 + 1;
                            t1.x = mmmToPixels2(GetCellX(1, k, sheet));
                            t1.y = mmmToPixels2(GetCellY(pageInfo.startRow, i, sheet));
                            t1.width = mmmToPixels2(cell.GetColWidth(unitType, k, 0));
                            t1.height = mmmToPixels2(cell.GetRowHeight(unitType, i, 0));
                            t = t1;
                        }
                        else
                        {
                            t1.col = k;
                            t1.row = i;                           
                            t1.rowSpan = 1;
                            t1.colSpan = 1;
                            t1.x = mmmToPixels2(GetCellX(1, k, sheet));
                            t1.y = mmmToPixels2(GetCellY(pageInfo.startRow, i, sheet));
                            t1.width = mmmToPixels2(cell.GetColWidth(unitType, k, 0));
                            t1.height = mmmToPixels2(cell.GetRowHeight(unitType, i, 0));
                            t1.isMergeCell = true;
                            t1.mergeTo = new CellInfo1()
                            {
                                col = t.col,
                                row = t.row
                            };
                            t1.cellPositionTotal = new CellPositionTotal()
                            {
                                x = t1.x,
                                y = t1.y,
                                width = t1.width,
                                height = t1.height
                            };
                        }                       
                        Model.CellBorder cellBorder = new Model.CellBorder();
                        if (i == r1)
                        {
                            cellBorder.top = cell.GetCellBorder(k, i, sheet, 1);
                        }
                        if (i == r2)
                        {
                            cellBorder.bottom = cell.GetCellBorder(k, i, sheet, 3);
                        }
                        if (k == c1)
                        {
                            cellBorder.left = cell.GetCellBorder(k, i, sheet, 0);
                        }
                        if (k == c2)
                        {
                            cellBorder.right = cell.GetCellBorder(k, i, sheet, 2);
                        }
                        cellBorder.col = k;
                        cellBorder.row = i;
                        cellBorderList.Add(cellBorder);
                        var cellPosition = new Model.CellPosition()
                        {
                            col = k,
                            row = i,
                            x = mmmToPixels2(GetCellX(1, k, sheet)),
                            y = mmmToPixels2(GetCellY(pageInfo.startRow, i, sheet)),
                            width = mmmToPixels2(cell.GetColWidth(unitType, k, 0)),
                            height = mmmToPixels2(cell.GetRowHeight(unitType, i, 0))
                        };
                        cellPositionList.Add(cellPosition);
                        t1.cellBorderList = new List<CellBorder>() { cellBorder };
                        t1.cellPostionList = new List<CellPosition> { cellPosition };
                        L.Add(t1);
                    }
                }
            }
            else
            {
                width = cell.GetColWidth(unitType, col, sheet);
                height = cell.GetRowHeight(unitType, row, sheet);
                cellBorderList.Add(new Model.CellBorder()
                {
                    col = col,
                    row = row,
                    left = cell.GetCellBorder(col, row, sheet, 0),
                    top = cell.GetCellBorder(col, row, sheet, 1),
                    right = cell.GetCellBorder(col, row, sheet, 2),
                    bottom = cell.GetCellBorder(col, row, sheet, 3),
                    leftColor = cell.GetColor(cell.GetCellBorderClr(col, row, sheet, 0)),
                    topColor = cell.GetColor(cell.GetCellBorderClr(col, row, sheet, 1)),
                    rightColor = cell.GetColor(cell.GetCellBorderClr(col, row, sheet, 2)),
                    bottomColor = cell.GetColor(cell.GetCellBorderClr(col, row, sheet, 3)),
                });
                cellPositionList.Add(new Model.CellPosition()
                {
                    col = col,
                    row = row,
                    x = mmmToPixels2(GetCellX(1, col, sheet)),
                    y = mmmToPixels2(GetCellY(pageInfo.startRow, row, sheet)),
                    width = mmmToPixels2(cell.GetColWidth(unitType, col, 0)),
                    height = mmmToPixels2(cell.GetRowHeight(unitType, row, 0))
                });
                haveDisposeedCellList.Add(new MyPoint()
                {
                    X = col,
                    Y = row
                });
                CellInfo t1 = new CellInfo();
                t1.isMergeCell = false;
                t1.col = col;
                t1.row = row;                
                t1.rowSpan = 1;
                t1.colSpan = 1;
                t1.x = mmmToPixels2(GetCellX(1, col, sheet));
                t1.y = mmmToPixels2(GetCellY(pageInfo.startRow, row, sheet));
                t1.width = width;
                t1.height = height;
                L.Add(t1);
            }
        }

        private double mmmToPixels2(double v)
        {
            if (unitType == 1) //不作转换
            {
                return v;
            }
            return (72 * (1 / 25.4f)) * (v / 10);
        }

        private double GetCellY(int startRow, int row, int sheet)
        {
            var y = yPy;
            for (var i = startRow; i < row; i++)
            {
                y += cell.GetRowHeight(unitType, i, sheet);
            }
            return y;
        }

        private double GetCellX(int startCol, int col, int sheet)
        {
            var x = xPy;
            for (var i = startCol; i < col; i++)
            {
                x += cell.GetColWidth(unitType, i, sheet);
            }
            return x;
        }
    }
}
