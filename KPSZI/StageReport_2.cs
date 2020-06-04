using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;

namespace KPSZI
{
    abstract partial class StageReport
    {
        protected class HtmlElement
        {
            public string Text { get; protected set; }
            protected Color ForeColor { get; set; }
            public byte Bold { get; protected set; }
            public Word.WdColor wdForeColor { get; protected set; }

            public HtmlElement(string t, string fColor = "000000", byte bold = 0)
            {
                Text = t;
                Bold = bold;
                ForeColor = GetColor(fColor);
                wdForeColor = (Word.WdColor)(ForeColor.R + 0x100 * ForeColor.G + 0x10000 * ForeColor.B);
            }
            protected Color GetColor(string color)
            {
                string red = color.Substring(0, 2);
                string green = color.Substring(2, 2);
                string blue = color.Substring(4, 2);

                int rCode = int.Parse(red, System.Globalization.NumberStyles.HexNumber);
                int gCode = int.Parse(green, System.Globalization.NumberStyles.HexNumber);
                int bCode = int.Parse(blue, System.Globalization.NumberStyles.HexNumber);

                return Color.FromArgb(rCode, gCode, bCode);
            }
        }

        protected class HtmlTableElement : HtmlElement
        {
            
            public Word.WdColor wdBackColor { get; private set; }
            private Color BackColor { get; set; }
            public int Rowspan { get; private set; }
            public int Colspan { get; private set; }

            public HtmlTableElement(string t, string bColor = "FFFFFF", string fColor = "000000", byte bold = 0, int rspan = 1, int cspan = 1)
                :base(t, fColor, bold)
            {                
                BackColor = GetColor(bColor);                                
                wdBackColor = (Word.WdColor)(BackColor.R + 0x100 * BackColor.G + 0x10000 * BackColor.B);                
                Rowspan = rspan;
                Colspan = cspan;
            }        

        }
    }
}
