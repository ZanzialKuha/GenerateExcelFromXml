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

            // Нельзя повторно инициализировать наборы стилей
            WorkbookStylesPart StylesCheck = Document.WorkbookPart.WorkbookStylesPart;
            if (StylesCheck != null)
            {
                return;
            }

            WorkbookStylesPart workbookStylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();

            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            // набор шрифтов
            Fonts fonts1 = new Fonts() { Count = (UInt32Value)0U, KnownFonts = true };

            // В коде Excel заливки с Id 0 и 1 строго зашиты как "None" и "Gray125", поэтому для корректной работы исключим эти Id
            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            CellFill CellFill0 = new CellFill() { PatternType = PatternValues.None };
            Fill fill1 = CellFill0.Fill;

            CellFill CellFill1 = new CellFill() { PatternType = PatternValues.Gray125 };
            Fill fill2 = CellFill1.Fill;

            fills1.Append(fill1);
            fills1.Append(fill2);

            // набор границ
            Borders borders1 = new Borders() { Count = (UInt32Value)0U };

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            // форматы ячеек (шрифт + заливка + границы)
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)0U };

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
