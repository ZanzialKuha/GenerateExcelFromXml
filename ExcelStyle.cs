using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    public class CellFont
    {
        public UInt32Value Id { get; internal set; }
        public Font Font { get; } = new Font();
        public double FontSize { get; set; } = 10D;
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public bool Underline { get; set; } = false;
        public bool Strike { get; set; } = false;
        public UInt32Value ColorTheme { get; set; } = 1U;
        public string FontName { get; set; } = "Calibri";
        public Int32Value FontFamilyNumbering { get; set; } = 2;
        public EnumValue<FontSchemeValues> FontScheme { get; set; } = FontSchemeValues.Minor;
        /* Набор символов:
         - Arabic = 178,
         - Baltic = 186,
         - ChineseBig5 = 136,
         - ChineseGB2312 = 134,
         - EastEurope = 238,
         - Greek = 161,
         - Hebrew = 177,
         - JapaneseShiftJIS = 128,
         - KoreanHangeul = 129,
         - KoreanJohab = 130,
         - Russian = 204,
         - Thai = 222,
         - Turkish = 162,
         - Vietnamese = 163,
         - Symbol = 2,
         - ANSI = 0,
         - MAC = 77,
         - OEM = 255
        */
        public int FontCharSet { get; set; } = 204;

        public void Append()
        {
            FontSize fontSize1 = new FontSize() { Val = FontSize };
            Color color1 = new Color() { Theme = ColorTheme };
            FontName fontName1 = new FontName() { Val = FontName };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = FontFamilyNumbering };
            FontScheme fontScheme1 = new FontScheme() { Val = FontScheme };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = FontCharSet };

            Font.Append(fontSize1);
            Font.Append(color1);
            Font.Append(fontName1);
            Font.Append(fontFamilyNumbering1);
            Font.Append(fontScheme1);
            Font.Append(fontCharSet2);

            if (Bold)
            {
                Bold bold1 = new Bold();
                Font.Append(bold1);
            }

            if (Italic)
            {
                Italic italic1 = new Italic();
                Font.Append(italic1);
            }

            if (Underline)
            {
                Underline underline1 = new Underline();
                Font.Append(underline1);
            }

            if (Strike)
            {
                Strike strike1 = new Strike();
                Font.Append(strike1);
            }
        }
    }

    public class CellFill
    {
        public UInt32Value Id { get; internal set; }
        public Fill Fill { get; } = new Fill();
        public EnumValue<PatternValues> PatternType { get; set; } = PatternValues.None;
        public HexBinaryValue PatternBackgroundRgb { get; set; } = "";
        public HexBinaryValue PatternForegroundRgb { get; set; } = "";
        public void Append()
        {
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternType };

            if (PatternBackgroundRgb != "")
            {
                patternFill1.Append(new BackgroundColor() { Rgb = "FF" + PatternBackgroundRgb });
            }
            if (PatternForegroundRgb != "")
            {
                patternFill1.Append(new ForegroundColor() { Rgb = "FF" + PatternForegroundRgb });
            }

            Fill.Append(patternFill1);
        }
    }

    public class CellBorder
    {
        public UInt32Value Id { get; internal set; }
        public Border Border { get; } = new Border();
        public EnumValue<BorderStyleValues> LeftStyle { get; set; } = BorderStyleValues.None;
        public UInt32Value LeftColor { get; set; } = 0;
        public EnumValue<BorderStyleValues> RightStyle { get; set; } = BorderStyleValues.None;
        public UInt32Value RightColor { get; set; } = 0;
        public EnumValue<BorderStyleValues> TopStyle { get; set; } = BorderStyleValues.None;
        public UInt32Value TopColor { get; set; } = 0;
        public EnumValue<BorderStyleValues> BottomStyle { get; set; } = BorderStyleValues.None;
        public UInt32Value BottomColor { get; set; } = 0;
        public EnumValue<BorderStyleValues> DiagonalStyle { get; set; } = BorderStyleValues.None;
        public UInt32Value DiagonalColor { get; set; } = 0;

        public void Append()
        {
            LeftBorder LeftBorder = new LeftBorder() { Style = LeftStyle, Color = new Color() { Indexed = LeftColor } };
            Border.Append(LeftBorder);

            RightBorder RightBorder = new RightBorder() { Style = RightStyle, Color = new Color() { Indexed = RightColor } };
            Border.Append(RightBorder);

            TopBorder TopBorder = new TopBorder() { Style = TopStyle, Color = new Color() { Indexed = TopColor } };
            Border.Append(TopBorder);

            BottomBorder BottomBorder = new BottomBorder() { Style = BottomStyle, Color = new Color() { Indexed = BottomColor } };
            Border.Append(BottomBorder);

            DiagonalBorder DiagonalBorder = new DiagonalBorder() { Style = DiagonalStyle, Color = new Color() { Indexed = DiagonalColor } };
            Border.Append(DiagonalBorder);
        }
    }

    public class CellAlignment
    {
        public Alignment Alignment { get; } = new Alignment();
        public EnumValue<HorizontalAlignmentValues> Horizontal { get; set; } = HorizontalAlignmentValues.Left;
        public EnumValue<VerticalAlignmentValues> Vertical { get; set; } = VerticalAlignmentValues.Top;
        public bool WrapText { get; set; } = false;
        public void Append()
        {
            Alignment.Horizontal = Horizontal;
            Alignment.Vertical = Vertical;
            Alignment.WrapText = WrapText;
        }
    }
    public class ExcelStyle
    {
        public UInt32Value Id { get; internal set; }
        public CellFont CellFont { get; set; } = new CellFont();
        public CellFill CellFill { get; set; } = new CellFill();
        public CellBorder CellBorder { get; set; } = new CellBorder();
        public CellAlignment CellAlignment { get; set; } = new CellAlignment();

        public ExcelStyle()
        {

        }
    }
}
