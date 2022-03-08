using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentXML = DocumentFormat.OpenXml.ExtendedProperties;
using VariantTypes = DocumentFormat.OpenXml.VariantTypes;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    /// <summary>
    /// Вспомогательный класс для создания набора стилей (WorkbookStylesPart) документа (см. документацию Open XML)
    /// </summary>
    public class CommonWorkbookStylesPart : IWorkbookStylesPart
    {
        /// <summary>
        /// Генерация набора стилей (WorkbookStylesPart) (см. документацию Open XML)
        /// </summary>
        public virtual void GenerateWorkbookStylesPart(SpreadsheetDocument Document)
        {
            WorkbookPart WorkbookPart = (WorkbookPart)Document.WorkbookPart;
            WorkbookStylesPart workbookStylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();

            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            // набор шрифтов
            Fonts fonts1 = new Fonts() { Count = (UInt32Value)4U, KnownFonts = true };

            // Calibri, 11пт, обычный
            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            // Calibri, 11пт, жирный
            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet1);
            font2.Append(fontScheme2);

            // Calibri, 10пт, обычный
            Font font3 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 10D };
            Color color3 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontCharSet2);
            font3.Append(fontScheme3);

            // Calibri, 10пт, жирный
            Font font4 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = 10D };
            Color color4 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(bold3);
            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontCharSet3);
            font4.Append(fontScheme4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            // набор заливок
            Fills fills1 = new Fills() { Count = (UInt32Value)1U };

            // без заливки
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            fills1.Append(fill1);

            // набор границ
            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            // границы отсутствуют
            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            // тонкие границы со всех сторон
            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color6);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color7 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color7);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color8);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color9 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color9);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            // форматы ячеек (шрифт + заливка + границы)
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat3.Append(alignment1);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };
            cellFormat4.Append(alignment2);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };
            cellFormat5.Append(alignment3);

            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            cellFormats.Append(cellFormat5);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

            stylesheet.Append(fonts1);
            stylesheet.Append(fills1);
            stylesheet.Append(borders1);
            stylesheet.Append(cellStyleFormats1);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles1);
            stylesheet.Append(differentialFormats1);
            stylesheet.Append(tableStyles1);

            workbookStylesPart.Stylesheet = stylesheet;
        }
    }
}
