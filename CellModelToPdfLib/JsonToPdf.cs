using CellModelToPdfLib.Model;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CellModelToPdfLib
{
    public class JsonToPdf
    {
        CellJson cellJson = null;
        List<PageInfo> pageInfos = null;
        PdfDocument document = null;
        PdfPage page = null;
        XGraphics xGraphics = null;
        string fontPath = "";
        List<XFontDrawCache> xFontDrawCaches = new List<XFontDrawCache>();
        List<MeasureTextCache> measureTextCaches = new List<MeasureTextCache>();

        public JsonToPdf(CellJson cellJson, string _fontPath)
        {
            GlobalV.line1w = Comman.mmmToPixels2(2.0f);
            GlobalV.line2w = Comman.mmmToPixels2(3.8f) + 1;
            GlobalV.line3w = Comman.mmmToPixels2(5.0f) + 1;
            this.cellJson = cellJson;
            fontPath = _fontPath;
            pageInfos = cellJson.pages;
            document = new PdfDocument();
        }

        public void start(string saveToFile = "")
        {
            double allReadyDrawHeight = 0;
            foreach (var pageInfo in pageInfos)
            {
                page = document.AddPage();
                page.Width = pageInfo.paperWidth;
                page.Height = pageInfo.paperHeight;
                xGraphics = XGraphics.FromPdfPage(page);
                绘制背景图(pageInfo);
                var cells = cellJson.cells.FindAll(p => p.isMergeCell == false && pageInfo.startRow <= p.row && p.row <= pageInfo.endRow);
                foreach (var cellInfo in cells)
                {
                    绘制区域(cellInfo);
                    绘制文本(cellInfo);
                    绘制图片(pageInfo, cellInfo);
                }
                绘制浮动图片(pageInfo, allReadyDrawHeight);
                allReadyDrawHeight += pageInfo.contentHeight;
            }
            if (!string.IsNullOrEmpty(saveToFile))
            {
                document.Save(saveToFile);
            }
            else
            {
                document.Save("test.pdf");
            }
        }

        private void 绘制浮动图片(PageInfo pageInfo, double allReadyDrawHeight)
        {
            if (cellJson.floatImages.Count == 0)
            {
                return;
            }
            var L = cellJson.floatImages.FindAll(p => p.ypos >= allReadyDrawHeight && p.ypos <= pageInfo.contentHeight + allReadyDrawHeight);
            if (L == null || L.Count == 0)
            {
                return;
            }
            foreach (var floatImage in L)
            {
                var index = floatImage.index;
                var o = cellJson.images.Find(p => p.index == index);
                if (o == null)
                {
                    return;
                }
                float opacityYouWant = 0.6f; // opacityYouWant has to be a value between 0.0 and 1.0
                Image myTransparentImage = SetImageOpacity(Image.FromStream(new MemoryStream(o.image)), opacityYouWant);
                MemoryStream memoryStream = new MemoryStream();
                myTransparentImage.Save(memoryStream, ImageFormat.Png);
                XImage image = XImage.FromStream(memoryStream);
                xGraphics.DrawImage(image, new XRect(
                    floatImage.xpos + pageInfo.marginLeft,
                    floatImage.ypos + pageInfo.marginTop - allReadyDrawHeight,
                    floatImage.width,
                    floatImage.height)
                    );
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="image"></param>
        /// <param name="opacity">opacityYouWant has to be a value between 0.0 and 1.0</param>
        /// <returns></returns>
        private Image SetImageOpacity(Image image, float opacity)
        {
            //create a Bitmap the size of the image provided  
            Bitmap bmp = new Bitmap(image.Width, image.Height);

            //create a graphics object from the image  
            using (Graphics gfx = Graphics.FromImage(bmp))
            {

                //create a color matrix object  
                ColorMatrix matrix = new ColorMatrix();

                //set the opacity  
                matrix.Matrix33 = opacity;

                //create image attributes  
                ImageAttributes attributes = new ImageAttributes();

                //set the color(opacity) of the image  
                attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                //now draw the image  
                gfx.DrawImage(image, new Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
            }
            return bmp;
        }


        private Image SetImageOpacity2(Image image)
        {
            //create a Bitmap the size of the image provided  
            Bitmap bmp = new Bitmap(image.Width, image.Height);

            //create a graphics object from the image  
            using (Graphics gfx = Graphics.FromImage(bmp))
            {
                //create image attributes  
                ImageAttributes attributes = new ImageAttributes();

                attributes.SetColorKey(Color.FromArgb(127, 127, 127), Color.FromArgb(255, 255, 255), ColorAdjustType.Default);

                //now draw the image  
                gfx.DrawImage(image, new Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
            }
            return bmp;
        }


        private void 绘制背景图(PageInfo pageInfo)
        {
            if (cellJson.backGroundImageInfo == null)
            {
                return;
            }
            var index = cellJson.backGroundImageInfo.index;
            var style = cellJson.backGroundImageInfo.style;
            var o = cellJson.images.Find(p => p.index == index);
            if (o == null)
            {
                return;
            }
            var image = XImage.FromStream(new MemoryStream(o.image));
            XRect xRect = new XRect(
                pageInfo.marginLeft,
                pageInfo.marginTop,
                page.Width - pageInfo.marginLeft - pageInfo.marginRight,
                page.Height - pageInfo.marginTop - pageInfo.marginBottom
                );
            if (style == 0) //平铺
            {
                平铺绘制图片(image, xRect);
            }
            else if (style == 1) //居中
            {
                居中绘制图片(image, xRect, 2, 2);
            }
            else if (style == 2) //拉伸
            {
                拉伸绘制图片(image, xRect);
            }
            else
            {
                居中绘制图片(image, xRect, 2, 2);
            }
        }

        private void 拉伸绘制图片(XImage image, XRect xRect)
        {
            拉伸绘制图片(image, xRect.X, xRect.Y, xRect.Width, xRect.Height);
        }

        private void 绘制图片(PageInfo pageInfo, CellInfo cellInfo)
        {
            if (cellInfo.cellImage == null)
            {
                return;
            }
            var index = cellInfo.cellImage.index;
            var style = cellInfo.cellImage.style;
            var halign = cellInfo.cellImage.halign;
            var valign = cellInfo.cellImage.valign;
            var o = cellJson.images.Find(p => p.index == index);
            if (o == null)
            {
                return;
            }
            var t2 = cellInfo.cellPositionTotal;
            XRect rect = GetRect(t2, cellInfo.stringFormat.fontSize);
            var t = new CellPosition()
            {
                x = rect.X,
                y = rect.Y,
                width = rect.Width,
                height = rect.Height
            };
            if (t.width <= 0 || t.height <= 0)
            {
                return;
            }
            var myTransparentImage = Image.FromStream(new MemoryStream(o.image));
            myTransparentImage = SetImageOpacity2(myTransparentImage);
            MemoryStream memoryStream = new MemoryStream();
            myTransparentImage.Save(memoryStream, ImageFormat.Png);
            XImage image = XImage.FromStream(memoryStream);
            if (style == 1)
            {
                拉伸绘制图片(image, t.x, t.y, t.width, t.height);
            }
            else if (style == 1 + 2)
            {
                居中绘制图片(image, new XRect(t.x, t.y, t.width, t.height), halign, valign);
            }
            else if (style == 4) //平铺
            {
                平铺绘制图片(image, new XRect(t.x, t.y, t.width, t.height));
            }
        }

        private void 拉伸绘制图片(XImage image, double x, double y, double width, double height)
        {
            xGraphics.DrawImage(image, x, y, width, height);
        }


        private void 居中绘制图片(XImage image, XRect xRect, int halign, int valign)
        {
            var w = (double)(image.PixelWidth * 72 / image.HorizontalResolution);
            var h = (double)(image.PixelHeight * 72 / image.VerticalResolution);
            var r1 = xRect.Width / w;
            var r2 = xRect.Height / h;
            var r = Math.Min(r1, r2);
            var w1 = w * r;
            var h1 = h * r;
            var x = xRect.X;
            var y = xRect.Y;
            if (halign == 0 || halign == 1)
            {
                //
            }
            else if (halign == 2)
            {
                x += (xRect.Width - w1) / 2;
            }
            else if (halign == 3)
            {
                x += xRect.Width - w1;
            }
            if (valign == 0 || valign == 1)
            {
                //
            }
            else if (valign == 2)
            {
                y += (xRect.Height - h1) / 2;
            }
            else if (valign == 3)
            {
                y += xRect.Height - h1;
            }
            xGraphics.DrawImage(image, x, y, w1, h1);
        }

        private void 平铺绘制图片(XImage image, XRect xRect)
        {
            var state = xGraphics.Save();
            xGraphics.IntersectClip(xRect);
            var x = xRect.X;
            var y = xRect.Y;
            var w = (double)(image.PixelWidth * 72 / image.HorizontalResolution);
            var h = (double)(image.PixelHeight * 72 / image.VerticalResolution);
            var w1 = 0d;
            var h1 = 0d;
            while (w1 < xRect.Width)
            {
                xGraphics.DrawImage(image, x, y, w, h);
                y = xRect.Y + h;
                h1 = h;
                while (h1 < xRect.Height)
                {
                    xGraphics.DrawImage(image, x, y, w, h);
                    y += h;
                    h1 += h;
                }
                x += w;
                y = xRect.Y;
                w1 += w;
            }
            xGraphics.Restore(state);
        }

        private void 绘制文本(CellInfo cellInfo)
        {
            if (string.IsNullOrEmpty(cellInfo.str))
            {
                return;
            }
            if (cellInfo.isHidden)
            {
                return;
            }
            if (cellInfo.cellPositionTotal == null)
            {
                return;
            }
            var t1 = cellInfo.stringFormat;
            var align = GetTextAlign(cellInfo.cellAlign);
            var rect = GetRect(cellInfo.cellPositionTotal, t1.fontSize);
            if (rect.Width == 0 || rect.Height == 0)
            {
                return;
            }
            var fontStyle = GetFontStyle2(t1.fontStyle);
            string str = cellInfo.str;
            if (str == "检\r\n测\r\n信息")
            {
                str = "检\r\n测\r\n信\r\n息";
            }
            List<string> strList = str.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList();
            double fontSize1 = t1.fontSize;
            double lineSpace = t1.lineSpace;
            var font = GetXFont(t1.fontFamily, fontSize1, GetFontStyle(t1.fontStyle));
            var isNeedAdjust = true;//GetIsNeedAdjust(strList, font, rect);
            if (isNeedAdjust)
            {
                bool isNeedAdjustFontSize = true;
                List<string> strListOriginal = Comman.DeepCopyList(strList);
                while (isNeedAdjustFontSize)
                {
                    font = GetXFont(t1.fontFamily, fontSize1, GetFontStyle(t1.fontStyle));
                    strList = Comman.DeepCopyList(strListOriginal);
                    strList = AdjustStrAr(strList, cellInfo.cellPositionTotal, font, lineSpace, ref fontSize1, ref isNeedAdjustFontSize);
                }
            }
            rect = GetRect(cellInfo.cellPositionTotal, t1.fontSize);
            if (rect.Width == 0 || rect.Height == 0)
            {
                return;
            }
            var t = new CellPosition()
            {
                x = rect.X,
                y = rect.Y,
                width = rect.Width,
                height = rect.Height
            };
            var textSize = MeasureText("我", font);
            var padingLeft = 0;
            var paddingTop = 0;
            XBrush xBrush = new XSolidBrush(Comman.ToXColor(cellInfo.stringFormat.fontColor));
            if (strList.Count > 1)
            {
                var y = GetY(rect, cellInfo.cellAlign, strList.Count, textSize, lineSpace);
                var height = textSize.Height;
                for (var i = 0; i < strList.Count; i++)
                {
                    List<SupInfo> supInfos = new List<SupInfo>();
                    List<SubInfo> subInfos = new List<SubInfo>();
                    string line = strList[i];
                    GetSupSubInfoList(ref line, ref supInfos, ref subInfos);
                    var isHaveSupSubInfos = (supInfos.Count > 0 || subInfos.Count > 0);
                    if (isHaveSupSubInfos)
                    {
                        int index = 0;
                        var supSubFont = GetXFont(t1.fontFamily, (double)(fontSize1 * 0.7), GetFontStyle(t1.fontStyle));
                        DrawLineHaveSupSubs(line, supInfos, subInfos, index, font, t, padingLeft, paddingTop, align,
                            height, y, supSubFont, textSize, cellInfo.cellAlign, xBrush);
                    }
                    else
                    {
                        xGraphics.DrawString(line, font, xBrush, new XRect(t.x + padingLeft, y + paddingTop, t.width, height), align);
                    }
                    y += (double)textSize.Height + lineSpace;
                }
            }
            else
            {
                List<SupInfo> supInfos = new List<SupInfo>();
                List<SubInfo> subInfos = new List<SubInfo>();
                string line = strList[0];
                GetSupSubInfoList(ref line, ref supInfos, ref subInfos);
                var isHaveSupSubInfos = (supInfos.Count > 0 || subInfos.Count > 0);
                if (isHaveSupSubInfos)
                {
                    var supSubFont = GetXFont(t1.fontFamily, (double)(fontSize1 * 0.7), GetFontStyle(t1.fontStyle));
                    DrawLineHaveSupSubs(line, supInfos, subInfos, 0, font, t, padingLeft, paddingTop, align,
                        t.height, t.y, supSubFont, textSize, cellInfo.cellAlign, xBrush);
                }
                else
                {
                    xGraphics.DrawString(line, font, xBrush, new XRect(t.x + padingLeft, t.y + paddingTop, t.width, t.height), align);
                }
            }
        }

        private XRect GetRect(CellPositionTotal t2, double fontSize)
        {
            var hmargin = 0;
            var vmargin = 0;
            Comman.GetHVMargin((int)fontSize, ref hmargin, ref vmargin);
            var t = new CellPosition()
            {
                x = t2.x + hmargin,
                y = t2.y + vmargin,
                width = t2.width - 2 * hmargin,
                height = t2.height - 2 * vmargin
            };
            if (t.width < 0)
            {
                t.width = 0;
            }
            if (t.height < 0)
            {
                t.height = 0;
            }
            return new XRect(t.x, t.y, t.width, t.height);
        }

        private void DrawLineHaveSupSubs(string line, List<SupInfo> supInfos, List<SubInfo> subInfos, int index, XFont font,
            CellPosition t, double padingLeft, double paddingTop, XStringFormat align, double height, double y, XFont supSubFont,
            XSize textSize, int cellAlign, XBrush xBrush)
        {
            List<MyText> L = GetStrList(line, supInfos, subInfos, index);
            DrawTextHaveSupSubsFromLeftToRight(L, t, font, supSubFont, padingLeft, paddingTop, y, height, align, cellAlign, textSize, xBrush, line);
        }

        private void DrawTextHaveSupSubsFromLeftToRight(List<MyText> L, CellPosition t, XFont font, XFont supSubFont,
            double padingLeft, double paddingTop, double y, double height, XStringFormat align, int cellAlign,
            XSize textSize, XBrush xBrush, string line)
        {
            var x = 0d;
            if (IsHaveCellAlignLeft(cellAlign))
            {
                x = t.x + padingLeft;
            }
            else if (IsHaveCellAlignCenter(cellAlign))
            {
                var t1 = MeasureText(line, font);
                x = t.x + (t.width - t1.Width) / 2;
            }
            else
            {
                var t1 = MeasureText(line, font);
                x = t.x + t.width - t1.Width;
            }
            for (var k = 0; k < L.Count; k++)
            {
                double width = MeasureText(L[k].text, font).Width;
                if (L[k].textType == Enums.MyTextType.Sup)
                {
                    xGraphics.DrawString(L[k].text, supSubFont, xBrush, new XRect(x, y + paddingTop - 3, width, height), align);
                }
                else if (L[k].textType == Enums.MyTextType.Sub)
                {
                    xGraphics.DrawString(L[k].text, supSubFont, xBrush, new XRect(x, y + paddingTop + 3, width, height), align);
                }
                else
                {
                    xGraphics.DrawString(L[k].text, font, xBrush, new XRect(x, y + paddingTop, width, height), align);
                }
                x += width;
            }
        }

        /// <summary>
        /// 单元格有水平居中对齐
        /// </summary>
        /// <param name="cellAlign"></param>
        /// <returns></returns>
        private bool IsHaveCellAlignCenter(int cellAlign)
        {
            return cellAlign == 4 || cellAlign == 4 + 8 || cellAlign == 4 + 16 || cellAlign == 4 + 32;
        }

        /// <summary>
        /// 单元格有水平居右对齐
        /// </summary>
        /// <param name="cellAlign"></param>
        /// <returns></returns>
        private bool IsHaveCellAlignRight(int cellAlign)
        {
            return cellAlign == 2 || cellAlign == 2 + 8 || cellAlign == 2 + 16 || cellAlign == 2 + 32;
        }

        /// <summary>
        /// 单元格有水平左对齐
        /// </summary>
        /// <param name="cellAlign"></param>
        /// <returns></returns>
        private bool IsHaveCellAlignLeft(int cellAlign)
        {
            return cellAlign == 0 || cellAlign == 1 || cellAlign == 1 + 8 || cellAlign == 1 + 16 || cellAlign == 1 + 32;
        }

        private List<MyText> GetStrList(string line, List<SupInfo> supInfos, List<SubInfo> subInfos, int index)
        {
            List<MyText> L = new List<MyText>();
            string s = "";
            for (var i = 0; i < line.Length; i++)
            {
                var o1 = supInfos.Find(p => p.start == i + index);
                var o2 = subInfos.Find(p => p.start == i + index);
                if (o1 != null)
                {
                    if (s != "")
                    {
                        L.Add(new MyText()
                        {
                            textType = Enums.MyTextType.Normal,
                            text = s
                        });
                        s = "";
                    }
                    L.Add(new MyText()
                    {
                        textType = Enums.MyTextType.Sup,
                        text = ""
                    });
                    int index1 = 0;
                    for (var k = i; k < line.Length; k++)
                    {
                        if (k + index > o1.end)
                        {
                            break;
                        }
                        L[L.Count - 1].text += line[k];
                        index1++;
                    }
                    i += index1 - 1;
                }
                else if (o2 != null)
                {
                    if (s != "")
                    {
                        L.Add(new MyText()
                        {
                            textType = Enums.MyTextType.Normal,
                            text = s
                        });
                        s = "";
                    }
                    L.Add(new MyText()
                    {
                        textType = Enums.MyTextType.Sub,
                        text = ""
                    });
                    int index1 = 0;
                    for (var k = i; k < line.Length; k++)
                    {
                        if (k + index > o2.end)
                        {
                            break;
                        }
                        L[L.Count - 1].text += line[k];
                        index1++;
                    }
                    i += index1 - 1;
                }
                else
                {
                    s += line[i].ToString();
                }
            }
            if (s != "")
            {
                L.Add(new MyText()
                {
                    textType = Enums.MyTextType.Normal,
                    text = s
                });
                s = "";
            }
            return L;
        }

        private void GetSupSubInfoList(ref string str, ref List<SupInfo> supInfos, ref List<SubInfo> subInfos)
        {
            if (!str.Contains("&End"))
            {
                return;
            }
            var supInfo = new SupInfo();
            var subInfo = new SubInfo();
            for (var i = 0; i < str.Length - 3; i++)
            {
                var t = str.Substring(i, 4);
                if (t == "&Sup")
                {
                    supInfo = new SupInfo() { start = i };
                    subInfo = null;
                    str = SubStr(str, i, 4);
                }
                else if (t == "&Sub")
                {
                    supInfo = null;
                    subInfo = new SubInfo() { start = i };
                    str = SubStr(str, i, 4);
                }
                else if (t == "&End")
                {
                    if (supInfo != null)
                    {
                        supInfo.end = i - 1;
                        supInfos.Add(supInfo);
                    }
                    else if (subInfo != null)
                    {
                        subInfo.end = i - 1;
                        subInfos.Add(subInfo);
                    }
                    supInfo = null;
                    subInfo = null;
                    str = SubStr(str, i, 4);
                    i = i - 1;
                }
            }
        }

        private string SubStr(string str, int pos1, int length)
        {
            if (pos1 == 0)
            {
                str = str.Substring(length);
            }
            else if (pos1 + length == str.Length)
            {
                str = str.Substring(0, pos1);
            }
            else
            {
                str = str.Substring(0, pos1) + str.Substring(pos1 + length);
            }
            return str;
        }

        private bool GetIsNeedAdjust(List<string> strList, XFont font, XRect rect)
        {
            foreach (var p in strList)
            {
                var t = MeasureText(p, font);
                if (t.Width > rect.Width || t.Height > rect.Height)
                {
                    return true;
                }
            }
            return false;
        }

        private List<string> AdjustStrAr(List<string> strList, CellPositionTotal t2, XFont font, double lineSpace, ref double fontSize, ref bool isNeedAdjustFontSize)
        {
            XRect rect = GetRect(t2, fontSize);
            if (rect.Width == 0 || rect.Height == 0)
            {
                return strList;
            }
            isNeedAdjustFontSize = false;
            List<SupInfo> supInfos = new List<SupInfo>();
            List<SubInfo> subInfos = new List<SubInfo>();
            for (var i = 0; i < strList.Count; i++)
            {
                var s = strList[i];
                GetSupSubInfoList(ref s, ref supInfos, ref subInfos); //这里只是为了排除上下标
                var t = MeasureText(s, font);
                if (t.Width > rect.Width)
                {
                    List<string> L1 = new List<string>();
                    List<string> L = new List<string>();
                    for (var k = 0; k < s.Length; k++)
                    {
                        L.Add(s[k].ToString());
                        t = MeasureText(string.Join("", L), font);
                        if (t.Width > rect.Width)
                        {
                            var 断句位置 = 得到断句位置(s, k);
                            if (断句位置 != -11111)
                            {
                                List<string> tempList = new List<string>();
                                if (断句位置 > 0)
                                {
                                    var t1 = L.Count - 1 - 断句位置;
                                    if (t1 < 0)
                                    {
                                        t1 = 0;
                                    }
                                    var t3 = 断句位置;
                                    if (t3 - t1 > L.Count)
                                    {
                                        t3 = L.Count - t1;
                                    }
                                    tempList = L.GetRange(t1, t3);
                                    for (var n = L.Count - 1; n > t1; n--)
                                    {
                                        L.RemoveAt(n);
                                    }
                                }
                                if (L.Count > 1)
                                {
                                    var t1 = string.Join("", L.GetRange(0, L.Count - 1));
                                    L1.Add(t1);
                                    L.Clear();
                                    L.Add(string.Join("", tempList) + s[k].ToString());
                                }
                                else
                                {
                                    L1.Add(string.Join("", L));
                                    L.Clear();
                                    fontSize--;
                                    if (fontSize > 3)
                                    {
                                        isNeedAdjustFontSize = true;
                                        return strList;
                                    }
                                }
                            }
                            else //缩小字体
                            {
                                fontSize--;
                                if (fontSize > 3)
                                {
                                    isNeedAdjustFontSize = true;
                                    return strList;
                                }
                            }
                        }
                    }
                    if (L.Count > 0)
                    {
                        L1.Add(string.Join("", L));
                    }
                    strList.InsertRange(i + 1, L1);
                    strList.RemoveAt(i);
                }
            }
            var textSize = MeasureText("我", font);
            if ((strList.Count * textSize.Height + (strList.Count - 1) * lineSpace) > rect.Height)
            {
                if (fontSize > 3)
                {
                    fontSize--;
                    isNeedAdjustFontSize = true;
                }
            }
            return strList;
        }

        private int 得到断句位置(string s, int k)
        {
            if (k == 0)
            {
                return -11111;
            }
            var t1 = s[k - 1];
            var t2 = s[k];
            if (GlobalV.不能分开的字符.Contains(t1) && GlobalV.不能分开的字符.Contains(t2))
            {
                int index = 0;
                for (var i = k - 1; i >= 1; i--)
                {
                    if (GlobalV.不能分开的字符.Contains(s[i]))
                        index++;
                    else
                        break;
                }
                return index;
            }
            if (GlobalV.数字.Contains(t1) && GlobalV.数字.Contains(t2))
            {
                int index = 0;
                for (var i = k - 1; i >= 1; i--)
                {
                    if (GlobalV.数字.Contains(s[i]))
                        index++;
                    else
                        break;
                }
                return index;
            }            
            if (GlobalV.canNotInTheEndChars.Contains(t1))
            {
                return 1;
            }
            if (GlobalV.canNotInTheFirstChars.Contains(t2))
            {
                return -11111;
            }
            return 0;
        }

        private List<string> AdjustStrArBak(List<string> strList, CellPositionTotal t2, XFont font, double lineSpace, ref double fontSize, ref bool isNeedAdjustFontSize)
        {
            XRect rect = GetRect(t2, fontSize);
            if (rect.Width == 0 || rect.Height == 0)
            {
                return strList;
            }
            isNeedAdjustFontSize = false;
            List<SupInfo> supInfos = new List<SupInfo>();
            List<SubInfo> subInfos = new List<SubInfo>();
            for (var i = 0; i < strList.Count; i++)
            {
                var s = strList[i];
                GetSupSubInfoList(ref s, ref supInfos, ref subInfos); //这里只是为了排除上下标
                var t = MeasureText(s, font);
                if (t.Width > rect.Width)
                {
                    List<string> L1 = new List<string>();
                    List<string> L = new List<string>();
                    for (var k = 0; k < s.Length; k++)
                    {
                        L.Add(s[k].ToString());
                        t = MeasureText(string.Join("", L), font);
                        if (t.Width > rect.Width)
                        {
                            if (是否可以分行(s, k))
                            {
                                if (L.Count > 1)
                                {
                                    var t1 = string.Join("", L.GetRange(0, L.Count - 1));
                                    L1.Add(t1);
                                    L.Clear();
                                    L.Add(s[k].ToString());
                                }
                                else
                                {
                                    L1.Add(string.Join("", L));
                                    L.Clear();
                                    fontSize--;
                                    if (fontSize > 3)
                                    {
                                        isNeedAdjustFontSize = true;
                                        return strList;
                                    }
                                }
                            }
                            else //缩小字体
                            {
                                fontSize--;
                                if (fontSize > 3)
                                {
                                    isNeedAdjustFontSize = true;
                                    return strList;
                                }
                            }
                        }
                    }
                    if (L.Count > 0)
                    {
                        L1.Add(string.Join("", L));
                    }
                    strList.InsertRange(i + 1, L1);
                    strList.RemoveAt(i);
                }
            }
            var textSize = MeasureText("我", font);
            if ((strList.Count * textSize.Height + (strList.Count - 1) * lineSpace) > rect.Height)
            {
                if (fontSize > 3)
                {
                    fontSize--;
                    isNeedAdjustFontSize = true;
                }
            }
            return strList;
        }

        private bool 是否可以分行(string s, int k)
        {
            if (k == 0) //只有一个字
            {
                return false;
            }
            var t1 = s[k - 1];
            var t2 = s[k];
            if (GlobalV.不能分开的字符.Contains(t1) && GlobalV.不能分开的字符.Contains(t2)) //数字不要分开
            {
                return false;
            }
            if (k - 2 >= 0)
            {
                var t3 = s[k - 2];
                if (t2 == ' ' &&
                    (
                        (t3 == '/' && t1 == 'T')
                        ||
                        (t3 == 'G' && t1 == 'B')
                        ||
                        (t3 == 'T' && t1 == 'B')
                        ||
                        (t3 == 'D' && t1 == 'B')
                    )
                    )
                {
                    return false;
                }
            }
            if (k - 2 >= 0 && k - 3 >= 0)
            {
                var t3 = s[k - 2];
                var t4 = s[k - 3];
                if (GlobalV.不能分开的字符.Contains(t2) && t1 == ' ' &&
                    (
                        (t4 == '/' && t3 == 'T')
                        ||
                        (t4 == 'G' && t3 == 'B')
                        ||
                        (t4 == 'T' && t3 == 'B')
                    )
                    )
                {
                    return false;
                }
            }
            if (k - 2 >= 0)
            {
                var t3 = s[k - 2];
                if (GlobalV.数字.Contains(t3) && t1 == '.' && GlobalV.数字.Contains(t2))
                {
                    return false;
                }
            }
            if (GlobalV.canNotInTheEndChars.Contains(t1))
            {
                return false;
            }
            if (GlobalV.canNotInTheFirstChars.Contains(t2))
            {
                return false;
            }
            return true;
        }

        private double GetY(XRect rect, int align, int length, XSize textSize, double lineSpace)
        {
            if (align == 1 || align == 2 || align == 4 || align == 8
                || align == 1 + 8 || align == 2 + 8 || align == 4 + 8)
            {
                return (double)rect.Y;
            }
            if (align == 0 || align == 32
                || align == 1 + 32 || align == 2 + 32 || align == 4 + 32)
            {
                var t1 = rect.Height - textSize.Height * length - lineSpace * (length - 1);
                if (t1 > 0)
                {
                    return (double)((t1 / 2) + rect.Y);
                }
                else
                {
                    return (double)rect.Y;
                }
            }
            else
            {
                var t1 = rect.Height - textSize.Height * length - lineSpace * (length - 1);
                if (t1 > 0)
                {
                    return (double)(t1 + rect.Y);
                }
                else
                {
                    return (double)rect.Y;
                }
            }
        }

        private XSize MeasureText(string text, XFont font)
        {
            var o = measureTextCaches.Find(p => p.text == text && p.font.Equals(font));
            if (o != null)
            {
                return o.size;
            }
            var size = xGraphics.MeasureString(text, font);
            measureTextCaches.Add(new MeasureTextCache()
            {
                text = text,
                font = font,
                size = size
            });
            return size;
        }

        private XFont GetXFont(string fontFamily, double fontSize, XFontStyle xFontStyle)
        {
            XFont font;
            string strFontPath = "";
            if (fontFamily == "Times New Roman")
            {
                fontFamily = "宋体";
            }
            if (fontFamily == "宋体")
            {
                if (xFontStyle == XFontStyle.Regular)
                {
                    var o = xFontDrawCaches.Find(p => p.fontFamily == fontFamily && p.fontSize == fontSize && p.xFontStyle == xFontStyle);
                    if (o != null)
                    {
                        return o.xFont;
                    }
                    System.Drawing.Text.PrivateFontCollection pfcFonts = new System.Drawing.Text.PrivateFontCollection();
                    strFontPath = Path.Combine(fontPath, "SimSun.ttf");//字体设置为宋体
                    pfcFonts.AddFontFile(strFontPath);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    font = new XFont(pfcFonts.Families[0], fontSize, xFontStyle, options); //只能用XFontStyle.Regular,不能用其它的
                }
                else
                {
                    fontFamily = "华文宋体";
                    var o = xFontDrawCaches.Find(p => p.fontFamily == fontFamily && p.fontSize == fontSize && p.xFontStyle == xFontStyle);
                    if (o != null)
                    {
                        return o.xFont;
                    }
                    System.Drawing.Text.PrivateFontCollection pfcFonts = new System.Drawing.Text.PrivateFontCollection();
                    strFontPath = Path.Combine(fontPath, "STSONG.TTF");//字体设置华文宋体
                    pfcFonts.AddFontFile(strFontPath);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    font = new XFont(pfcFonts.Families[0], fontSize, xFontStyle, options);
                }
            }
            else
            {
                try
                {
                    var o = xFontDrawCaches.Find(p => p.fontFamily == fontFamily && p.fontSize == fontSize && p.xFontStyle == xFontStyle);
                    if (o != null)
                    {
                        return o.xFont;
                    }
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    font = new XFont(fontFamily, fontSize, xFontStyle, options);
                }
                catch
                {
                    fontFamily = "华文宋体";
                    var o = xFontDrawCaches.Find(p => p.fontFamily == fontFamily && p.fontSize == fontSize && p.xFontStyle == xFontStyle);
                    if (o != null)
                    {
                        return o.xFont;
                    }
                    System.Drawing.Text.PrivateFontCollection pfcFonts = new System.Drawing.Text.PrivateFontCollection();
                    strFontPath = Path.Combine(fontPath, "STSONG.TTF");//字体设置华文宋体
                    pfcFonts.AddFontFile(strFontPath);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    font = new XFont(pfcFonts.Families[0], fontSize, xFontStyle, options);
                }
            }

            xFontDrawCaches.Add(new XFontDrawCache()
            {
                fontFamily = fontFamily,
                fontSize = fontSize,
                xFontStyle = xFontStyle,
                xFont = font
            });

            return font;
        }

        //private FontFamily GetFontFamily(string fontFamily)
        //{
        //    InstalledFontCollection MyFont = new InstalledFontCollection();
        //    FontFamily[] MyFontFamilies = MyFont.Families;
        //    foreach (var p in MyFontFamilies)
        //    {
        //        if (p.Name == fontFamily)
        //        {
        //            return p;
        //        }
        //    }
        //    throw new Exception("未找到安装的字体:" + fontFamily);
        //}

        private XStringFormat GetTextAlign(int cellAlign)
        {
            if (cellAlign == 0)
            {
                return XStringFormats.Center;
            }
            if (cellAlign == 1)
            {
                return XStringFormats.TopLeft;
            }
            if (cellAlign == 2)
            {
                return XStringFormats.TopRight;
            }
            if (cellAlign == 4)
            {
                return XStringFormats.TopCenter;
            }
            if (cellAlign == 8)
            {
                return XStringFormats.TopLeft;
            }
            if (cellAlign == 16)
            {
                return XStringFormats.BottomLeft;
            }
            if (cellAlign == 32)
            {
                return XStringFormats.CenterLeft;
            }
            if (cellAlign == 1 + 8)
            {
                return XStringFormats.TopLeft;
            }
            if (cellAlign == 2 + 8)
            {
                return XStringFormats.TopRight;
            }
            if (cellAlign == 4 + 8)
            {
                return XStringFormats.TopCenter;
            }
            if (cellAlign == 1 + 16)
            {
                return XStringFormats.BottomLeft;
            }
            if (cellAlign == 2 + 16)
            {
                return XStringFormats.BottomRight;
            }
            if (cellAlign == 4 + 16)
            {
                return XStringFormats.BottomCenter;
            }
            if (cellAlign == 1 + 32)
            {
                return XStringFormats.CenterLeft;
            }
            if (cellAlign == 2 + 32)
            {
                return XStringFormats.CenterRight;
            }
            if (cellAlign == 4 + 32)
            {
                return XStringFormats.Center;
            }
            throw new Exception("未知CellAlign:" + cellAlign);
        }

        private XFontStyle GetFontStyle(int fontStyle)
        {
            if (fontStyle == 0)
            {
                return XFontStyle.Regular;
            }
            if (fontStyle == 2)
            {
                return XFontStyle.Bold;
            }
            if (fontStyle == 4)
            {
                return XFontStyle.Italic;
            }
            if (fontStyle == 8)
            {
                return XFontStyle.Underline;
            }
            if (fontStyle == 16)
            {
                return XFontStyle.Strikeout;
            }
            if (fontStyle == 2 + 4)
            {
                return XFontStyle.BoldItalic;
            }
            return XFontStyle.Regular;
        }

        private FontStyle GetFontStyle2(int fontStyle)
        {
            if (fontStyle == 0)
            {
                return FontStyle.Regular;
            }
            if (fontStyle == 2)
            {
                return FontStyle.Bold;
            }
            if (fontStyle == 4)
            {
                return FontStyle.Italic;
            }
            if (fontStyle == 8)
            {
                return FontStyle.Underline;
            }
            if (fontStyle == 16)
            {
                return FontStyle.Strikeout;
            }
            if (fontStyle == 2 + 4)
            {
                return FontStyle.Bold | FontStyle.Italic;
            }
            return FontStyle.Regular;
        }

        private void 绘制区域(CellInfo cellInfo)
        {
            if (cellInfo.cellPostionList == null)
            {
                return;
            }
            for (var i = 0; i < cellInfo.cellPostionList.Count; i++)
            {
                var t1 = cellInfo.cellPostionList[i];
                var t2 = cellInfo.cellBorderList[i];
                if (t2.left > 1)
                {
                    DrawLine.画线(t2.left, Comman.ToXColor(t2.leftColor), xGraphics, new XPoint() { X = t1.x, Y = t1.y }, new XPoint() { X = t1.x, Y = t1.y + t1.height });
                }
                //else
                //{
                //    DrawLine.画线(6, Comman.ToXColor(t2.leftColor), xGraphics, new XPoint() { X = t1.x, Y = t1.y }, new XPoint() { X = t1.x, Y = t1.y + t1.height });
                //}

                if (t2.top > 1)
                {
                    DrawLine.画线(t2.top, Comman.ToXColor(t2.topColor), xGraphics, new XPoint() { X = t1.x, Y = t1.y }, new XPoint() { X = t1.x + t1.width, Y = t1.y });
                }
                //else
                //{
                //    DrawLine.画线(6, Comman.ToXColor(t2.topColor), xGraphics, new XPoint() { X = t1.x, Y = t1.y }, new XPoint() { X = t1.x + t1.width, Y = t1.y });
                //}

                if (t2.right > 1)
                {
                    DrawLine.画线(t2.right, Comman.ToXColor(t2.rightColor), xGraphics, new XPoint() { X = t1.x + t1.width, Y = t1.y }, new XPoint() { X = t1.x + t1.width, Y = t1.y + t1.height });
                }
                //else
                //{
                //    DrawLine.画线(6, Comman.ToXColor(t2.rightColor), xGraphics, new XPoint() { X = t1.x + t1.width, Y = t1.y }, new XPoint() { X = t1.x + t1.width, Y = t1.y + t1.height });
                //}

                if (t2.bottom > 1)
                {
                    DrawLine.画线(t2.bottom, Comman.ToXColor(t2.bottomColor), xGraphics, new XPoint() { X = t1.x, Y = t1.y + t1.height }, new XPoint() { X = t1.x + t1.width, Y = t1.y + t1.height });
                }
                //else
                //{
                //    DrawLine.画线(6, Comman.ToXColor(t2.bottomColor), xGraphics, new XPoint() { X = t1.x, Y = t1.y + t1.height }, new XPoint() { X = t1.x + t1.width, Y = t1.y + t1.height });
                //}
            }
        }
    }
}
