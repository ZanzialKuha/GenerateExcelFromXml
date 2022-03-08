using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentXML = DocumentFormat.OpenXml.ExtendedProperties;
using VariantTypes = DocumentFormat.OpenXml.VariantTypes;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    /// <summary>
    /// Вспомогательный класс для создания книги и листа (WorkbookPart и WorksheetPart) для документа (см. документацию Open XML)
    /// </summary>
    public class CommonWorkbookPart : IWorkbookPart
    {
        private WorksheetPart worksheetPart;
        /// <summary>
        /// Генерация книги и листа (WorkbookPart и WorksheetPart) (см. документацию Open XML)
        /// </summary>
        public virtual void GenerateWorkbookPart(string SheetName, SpreadsheetDocument Document)
        {
            // Нельзя создавать листы с одинаковыми названиями
            Sheet SheetCheck = Document.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => SheetName.Equals(s.Name));
            if (SheetCheck != null)
            {
                return;
            }

            // определим индекс листа
            Sheets Sheets = Document.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            UInt32Value SheetId = 1;
            if (Sheets.Elements<Sheet>().Count() > 0)
            {
                SheetId = Sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string relationshipId = String.Format("rId{0}", SheetId);

            // начнем создание листа с добавления узла WorksheetPart в родительский узел WorkbookPart, содержащий информацию по всем листам книги
            worksheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>(relationshipId);
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };
            SheetViews sheetViews1 = new SheetViews();
            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DefaultColumnWidth = 9.14D, DyDescent = 0.25D };

            SheetData sheetData1 = new SheetData();
            Columns columns1 = new Columns();

            // Columns не может инициализироваться пустым и/или инициализироваться после SheetData - фатальное разрушение структуры
            // Поэтому первый столбец создаем тут, все остальные - по мере обращения к ним
            Column CurrentColumn = new Column() { Min = 1U, Max = 1U, Width = 9.14D, CustomWidth = true };
            columns1.Append(CurrentColumn);

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);

            worksheetPart.Worksheet = worksheet1;
            worksheetPart.Worksheet.Save();

            // информация о "частях" книги
            DocumentXML.TitlesOfParts titlesOfParts1 = Document.ExtendedFilePropertiesPart.Properties.GetFirstChild<DocumentXML.TitlesOfParts>();
            VariantTypes.VTVector vTVector2;

            // добавим информацию о листе в книгу
            vTVector2 = titlesOfParts1.Elements<VariantTypes.VTVector>().First();
            vTVector2.Size++;
            
            VariantTypes.VTLPSTR vTLPSTR2 = new VariantTypes.VTLPSTR();
            vTLPSTR2.Text = SheetName;

            vTVector2.Append(vTLPSTR2);

            // создадим лист
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = SheetId, Name = SheetName };
            Sheets.Append(sheet);
        }
    }
}
