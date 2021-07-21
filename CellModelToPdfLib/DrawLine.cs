using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellModelToPdfLib
{
    public class DrawLine
    {
        public static void 画线(int style, XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            if ((xPoint2.X - xPoint1.X) == 0 && (xPoint2.Y - xPoint1.Y == 0))
            {
                return;
            }
            if (style == 2)
            {
                画细线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 3)
            {
                画中线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 4)
            {
                画粗线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 5)
            {
                画划线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 6)
            {
                画点线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 7)
            {
                画点划线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 8)
            {
                画点点划线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 9)
            {
                画粗划线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 10)
            {
                画粗点线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 11)
            {
                画粗点划线(xColor, xGraphics, xPoint1, xPoint2);
            }
            else if (style == 12)
            {
                画粗点点划线(xColor, xGraphics, xPoint1, xPoint2);
            }
        }

        private static void 画粗点点划线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.DashDotDot;
            xPen.Width = GlobalV.line3w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画粗点划线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.DashDot;
            xPen.Width = GlobalV.line3w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画粗点线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Dot;
            xPen.Width = GlobalV.line3w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画粗划线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Dash;
            xPen.Width = GlobalV.line3w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画点点划线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.DashDotDot;
            //xPen.Width = GlobalV.line1w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画点划线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.DashDot;
            //xPen.Width = GlobalV.line1w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画点线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Dot;
            //xPen.Width = GlobalV.line1w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画划线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Dash;
            //xPen.Width = GlobalV.line1w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画粗线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Solid;
            xPen.Width = GlobalV.line3w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画中线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Solid;
            xPen.Width = GlobalV.line2w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }

        private static void 画细线(XColor xColor, XGraphics xGraphics, XPoint xPoint1, XPoint xPoint2)
        {
            var xPen = new XPen(xColor);
            xPen.DashStyle = XDashStyle.Solid;
            //xPen.Width = GlobalV.line1w;
            xGraphics.DrawLine(xPen, xPoint1, xPoint2);
        }
    }
}
