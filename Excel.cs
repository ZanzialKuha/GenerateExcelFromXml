using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    public class Excel
    {
        private SpreadsheetDocument _document;
        private UInt32Value _colCount;
        private SheetData _SheetData;
        private MergeCells _MergeCells;
        private Worksheet _Worksheet;
        private WorksheetPart _WorksheetPart;
        private UInt32Value _SheetId;
        private string _ActiveSheetName;
        public string ActiveSheetName { get { return _ActiveSheetName; } }

        /// <summary>
        /// Создание документа для работы. Требуется передать количество столбцов для корректного заполнения шаблона
        /// </summary>
        public Excel(SpreadsheetDocument Document, UInt32Value ColCount)
        {
            _document = Document;
            _colCount = ColCount;
        }

        /// <summary>
        /// Создание новой книги Excel. Имя листа требуется при создании книги, т.к. Excel хранит информацию о листах в "заголовке".
        /// Для собственного переопределения процесса создания Excel файла используйте интерфейс IExtendedFilePropertiesPart.
        /// </summary>
        public void CreateExcel(string SheetName)
        {
            CreateExcel(SheetName, new CommonExtendedFilePropertiesPart());
            CreateSheet(SheetName);
            CreateStyles();
        }

        /// <summary>
        /// Создание новой книги Excel. Имя листа требуется при создании книги, т.к. Excel хранит информацию о листах в "заголовке".
        /// Для собственного переопределения процесса создания Excel файла используйте интерфейс IExtendedFilePropertiesPart.
        /// </summary>
        public void CreateExcel(string SheetName, IExtendedFilePropertiesPart SpreadsheetDocument)
        {
            SpreadsheetDocument.GenerateExtendedFilePropertiesPart(SheetName, _document);
        }

        /// <summary>
        /// Создание нового листа Excel.
        /// Для собственного переопределения создания листа используйте интерфейс IWorkbookPart вторым параметром
        /// </summary>
        public void CreateSheet(string SheetName)
        {
            CreateSheet(SheetName, new CommonWorkbookPart());
        }

        /// <summary>
        /// Создание нового листа Excel.
        /// Для собственного переопределения создания листа используйте интерфейс IWorkbookPart вторым параметром
        /// </summary>
        public void CreateSheet(string SheetName, IWorkbookPart WorkbookPart)
        {
            WorkbookPart.GenerateWorkbookPart(SheetName, _document);
            SetActiveSheet(SheetName);
        }

        /// <summary>
        /// Создание стилей
        /// Для собственного переопределения стилей используйте IWorkbookStylesPart в качестве параметра
        /// </summary>
        public void CreateStyles()
        {
            CreateStyles(new CommonWorkbookStylesPart());
        }

        /// <summary>
        /// Создание стилей
        /// Для собственного переопределения стилей используйте IWorkbookStylesPart в качестве параметра
        /// </summary>
        public void CreateStyles(IWorkbookStylesPart WorkbookStylesPart)
        {
            WorkbookStylesPart.GenerateWorkbookStylesPart(_document);

            CellFont CellFont0 = new CellFont() { FontSize = 11D, ColorTheme = 1U, FontName = "Calibri" };
            CellFont CellFont1 = new CellFont() { FontSize = 11D, ColorTheme = 1U, FontName = "Calibri", Bold = true };
            CellFont CellFont2 = new CellFont() { FontSize = 10D, ColorTheme = 1U, FontName = "Calibri", Bold = true };
            CellFont CellFont3 = new CellFont() { FontSize = 10D, ColorTheme = 1U, FontName = "Calibri", Bold = true };

            CellFill CellFill0 = new CellFill() { PatternType = PatternValues.None };

            CellBorder CellBorder0 = new CellBorder();
            CellBorder CellBorder1 = new CellBorder() { LeftColor = 64, LeftStyle = BorderStyleValues.Thin, RightColor = 64, RightStyle = BorderStyleValues.Thin, TopColor = 64, TopStyle = BorderStyleValues.Thin, BottomColor = 64, BottomStyle = BorderStyleValues.Thin };

            ExcelStyle ExcelStyle1 = new ExcelStyle() { CellFont = CellFont0, CellFill = CellFill0, CellBorder = CellBorder0 };
            ExcelStyle ExcelStyle2 = new ExcelStyle() { CellFont = CellFont1, CellFill = CellFill0, CellBorder = CellBorder0, CellAlignment = new CellAlignment() { Horizontal = HorizontalAlignmentValues.Center } };
            ExcelStyle ExcelStyle3 = new ExcelStyle() { CellFont = CellFont1, CellFill = CellFill0, CellBorder = CellBorder1, CellAlignment = new CellAlignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true } };
            ExcelStyle ExcelStyle4 = new ExcelStyle() { CellFont = CellFont0, CellFill = CellFill0, CellBorder = CellBorder1, CellAlignment = new CellAlignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true } };

            CreateOrFindCellStyle(ExcelStyle1);
            CreateOrFindCellStyle(ExcelStyle2);
            CreateOrFindCellStyle(ExcelStyle3);
            CreateOrFindCellStyle(ExcelStyle4);
        }

        /// <summary>
        /// Добавление данных в ячейку
        /// </summary>
        public void Append(string Position, UInt32Value Style = null, string Content = null, string MergeRangeStart = null, string MergeRangeEnd = null)
        {
            // найдем строку для позиции
            if (!String.IsNullOrEmpty(MergeRangeStart) && !String.IsNullOrEmpty(MergeRangeEnd))
            {
                MergeRange(MergeRangeStart, MergeRangeEnd, Style);
            }
            SetCell(Position, Style, Content);

            _Worksheet.Save();
        }

        /// <summary>
        /// Найти/создать ячейку для записи данных, установки стилей.
        /// </summary>
        private Cell SetCell(string reference, UInt32Value styleIndex = null, string value = null)
        {
            // найдем строку для позиции
            ExcelAddress Address = new ExcelAddress(reference);
            Row CurrentRow = GetRow(Address);
            Column CurrentColumn = GetColumn(Address);

            if (value == null)
            {
                value = "";
            }

            Cell cell = _Worksheet.Descendants<Cell>().Where(c => c.CellReference == Address.Address).FirstOrDefault();
            if (cell != null)
            {
                cell = _Worksheet.Descendants<Cell>().Where(c => c.CellReference == Address.Address).FirstOrDefault();

                if (styleIndex != null)
                {
                    cell.StyleIndex = styleIndex;
                }

                if (value.Length > 0)
                {
                    // добавим формулу, если первым символом передали "=" как в Excel
                    if (value[0] == '=')
                    {

                        cell.CellFormula = new CellFormula(value[1..]);
                    }
                    else
                    {
                        cell.CellValue = new CellValue(value);
                    }
                }
                else
                {
                    cell.CellValue = new CellValue(value);
                }

                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            else
            {
                cell = new Cell() { CellReference = Address.Address, DataType = CellValues.String };

                if (styleIndex != null)
                {
                    cell.StyleIndex = styleIndex;
                }

                CellValue cellValue = new CellValue();
                cellValue.Text = value;

                cell.Append(cellValue);

                CurrentRow.Append(cell);
            }

            return cell;
        }

        /// <summary>
        /// Найти/создать строку. При первом обращении к строке выполняется заполнение всех ячеек от 1 до _colCount для корректрой работы шаблона.
        /// </summary>
        private Row GetRow(ExcelAddress Position)
        {
            UInt32Value Row_Id = Convert.ToUInt32(Position.Row);
            Row CurrentRow;

            // попробуем найти уже созданную строку
            if (_SheetData.Elements<Row>().Where(r => r.RowIndex == Row_Id).Count() != 0)
            {
                CurrentRow = _SheetData.Elements<Row>().Where(r => r.RowIndex == Row_Id).First();
            }
            else
            {
                // создаем новую строку "по умолчанию"
                CurrentRow = new Row() { RowIndex = Row_Id, Spans = new ListValue<StringValue>() { InnerText = "1:100" }, CustomHeight = true, Height = 14.25D, DyDescent = 0.25D };
                _SheetData.Append(CurrentRow);

                for (UInt32Value CurrCol = 1; CurrCol <= _colCount; CurrCol++)
                {
                    Cell cell = new Cell() { CellReference = new ExcelAddress(String.Format("R{0}C{1}", Row_Id, CurrCol)).ToString(), DataType = CellValues.String };
                    CellValue cellValue = new CellValue();
                    cellValue.Text = "";
                    cell.Append(cellValue);
                    CurrentRow.Append(cell);
                }

            }

            return CurrentRow;
        }

        /// <summary>
        /// Найти/создать столбец. При первом обращении к столбцу выполняется создание данной колонки.
        /// </summary>
        private Column GetColumn(ExcelAddress Position)
        {
            UInt32Value Col_Id = Convert.ToUInt32(Position.Col);
            Column CurrentColumn;
            if (_SheetData.Elements<Column>().Where(r => r.Min == Col_Id).Count() != 0)
            {
                CurrentColumn = _SheetData.Elements<Column>().Where(r => r.Min == Col_Id).First();
            }
            else
            {
                Columns Columns = _Worksheet.Elements<Columns>().First();
                CurrentColumn = new Column() { Min = Col_Id, Max = Col_Id, Width = 9.14D, CustomWidth = true };
                Columns.Append(CurrentColumn);
            }

            return CurrentColumn;
        }

        /// <summary>
        /// Мерж диапазона ячеек. Для корректной работы границ обязательно надо инициализировать все ячейки внутри диапазона.
        /// </summary>
        private void MergeRange(string Start, string End, UInt32Value Style = null)
        {
            ExcelAddress StartAddress = new ExcelAddress(Start);
            ExcelAddress EndAddress = new ExcelAddress(End);

            UInt32Value StartRow = Convert.ToUInt32(StartAddress.Row);
            UInt32Value StartCol = Convert.ToUInt32(StartAddress.Col);

            UInt32Value EndRow = Convert.ToUInt32(EndAddress.Row);
            UInt32Value EndCol = Convert.ToUInt32(EndAddress.Col);

            for (UInt32Value CurrCol = StartCol; CurrCol <= EndCol; CurrCol++)
            {
                for (UInt32Value CurrRow = StartRow; CurrRow <= EndRow; CurrRow++)
                {
                    SetCell(String.Format("R{0}C{1}", CurrRow, CurrCol), Style);
                }
            }

            // нельзя инициализировать MergeCells без элементов - ошибка структуры xlsx
            _MergeCells = _Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (_MergeCells == null)
            {
                MergeCells MergeCells = new MergeCells();
                _Worksheet.Append(MergeCells);

                _MergeCells = _Worksheet.Elements<MergeCells>().FirstOrDefault();
            }

            _MergeCells.Append(new MergeCell() { Reference = String.Format("{0}:{1}", StartAddress.Address, EndAddress.Address) });
        }

        /// <summary>
        /// Установка ширины столбцов
        /// </summary>
        public void SetColumnWidth(UInt32Value StartCol, double[] Width)
        {
            for (UInt32Value CurrCol = StartCol; CurrCol <= Width.Length; CurrCol++)
            {
                Column CurrentColumn = GetColumn(new ExcelAddress(String.Format("R{0}C{1}", 1, CurrCol)));
                CurrentColumn.Width = (DoubleValue)Width[CurrCol - StartCol];
            }
        }

        /// <summary>
        /// Установка высоты строк
        /// </summary>
        public void SetRowHeight(UInt32Value StartRow, double[] Height)
        {
            for (UInt32Value CurrRow = StartRow; CurrRow <= Height.Length; CurrRow++)
            {
                Row CurrentRow = GetRow(new ExcelAddress(String.Format("R{0}C{1}", CurrRow, 1)));
                CurrentRow.Height = (DoubleValue)Height[CurrRow - StartRow];
            }
        }

        /// <summary>
        /// Получение данных ячеек
        /// </summary>
        public string GetCellData(string reference)
        {
            ExcelAddress Address = new ExcelAddress(reference);

            Cell cell = _Worksheet.Descendants<Cell>().Where(c => c.CellReference == Address.Address).FirstOrDefault();
            if (cell != null)
            {
                cell = _Worksheet.Descendants<Cell>().Where(c => c.CellReference == Address.Address).FirstOrDefault();
                return cell.CellValue.Text;
            }

            return String.Empty;
        }

        public void SetActiveSheet(string SheetName)
        {
            _ActiveSheetName = SheetName;
            WorkbookPart WorkbookPart = (WorkbookPart)_document.WorkbookPart;
            string relId = WorkbookPart.Workbook.Descendants<Sheet>().First(s => SheetName.Equals(s.Name)).Id;
            _SheetId = WorkbookPart.Workbook.Descendants<Sheet>().First(s => SheetName.Equals(s.Name)).SheetId;

            // данные должны быть уже заполнены при создании листа, обратимся к ним

            _WorksheetPart = (WorksheetPart)_document.WorkbookPart.GetPartById(relId);
            _Worksheet = _WorksheetPart.Worksheet;
            _SheetData = _Worksheet.Elements<SheetData>().First();

            WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            /*
            WorkbookView WorkbookView = _document.WorkbookPart.Workbook.BookViews.ChildElements.First<WorkbookView>();
            WorkbookView.ActiveTab = _SheetId;
*/
            /*
            var sheet = WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == SheetName);
            sheet.State = SheetStateValues.Hidden;
            */

        }

        /// <summary>
        /// Создание выпадающего списка значений по ссылке с другого диапазона ячеек и/или листов
        /// </summary>
        public void AddDropdownListLinkSheet(string RangeStart, string RangeEnd, string FromSheet, string TargetStart, string TargetEnd = "")
        {
            if (TargetEnd == "")
            {
                TargetEnd = TargetStart;
            }

            DataValidations DataValidations = _Worksheet.GetFirstChild<DataValidations>();
            if (DataValidations != null)
            {
                DataValidations.Count = DataValidations.Count + 1;
            }
            else
            {
                DataValidations = new DataValidations();
                DataValidations.Count = 1;
                _Worksheet.Append(DataValidations);
            }

            DataValidation DataValidation = new DataValidation
            {
                Type = DataValidationValues.List,
                AllowBlank = true,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Format("{0}:{1}", new ExcelAddress(TargetStart), new ExcelAddress(TargetEnd)) }
            };

            DataValidation.Append(
                new Formula1(string.Format("'{0}'!{1}:{2}", FromSheet, new ExcelAddress(RangeStart, true, true), new ExcelAddress(RangeEnd, true, true)))
                );
            DataValidations.Append(DataValidation);
        }

        /// <summary>
        /// Создание выпадающего списка значений в явном виде (ожидаемый разделитель значенией - символ ",")
        /// </summary>
        public void AddDropdownList(string DropdownString, string TargetStart, string TargetEnd = "")
        {
            if (TargetEnd == "")
            {
                TargetEnd = TargetStart;
            }

            DataValidations DataValidations = _Worksheet.GetFirstChild<DataValidations>();
            if (DataValidations != null)
            {
                DataValidations.Count = DataValidations.Count + 1;
            }
            else
            {
                DataValidations = new DataValidations();
                DataValidations.Count = 1;
                _Worksheet.Append(DataValidations);
            }

            DataValidation DataValidation = new DataValidation
            {
                Type = DataValidationValues.List,
                AllowBlank = true,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Format("{0}:{1}", new ExcelAddress(TargetStart), new ExcelAddress(TargetEnd)) }
            };

            DataValidation.Append(
                new Formula1 { Text = string.Format("\"{0}\"", DropdownString) }
                );
            DataValidations.Append(DataValidation);
        }

        /// <summary>
        /// WIP. Создание фильтра для поиска по столбцам
        /// </summary>
        private void AddFilter(string TargetStart, string TargetEnd = "")
        {
            Workbook Workbook = _document.WorkbookPart.Workbook;

            DefinedNames DefinedNames = Workbook.GetFirstChild<DefinedNames>();
            if (DefinedNames == null)
            {
                DefinedNames = new DefinedNames();
                Workbook.Append(DefinedNames);
            }

            if (TargetEnd == "")
            {
                TargetEnd = TargetStart;
            }

            DefinedName DefinedName = new DefinedName()
            {
                Name = "_xlnm._FilterDatabase",
                Text = string.Format("'{0}'!{1}:{2}", _ActiveSheetName, new ExcelAddress(TargetStart, true, true), new ExcelAddress(TargetEnd, true, true)),
                LocalSheetId = _SheetId
            };

            DefinedNames.Append(DefinedName);
        }

        public UInt32Value CreateOrFindCellStyle(ExcelStyle ExcelStyle)
        {
            WorkbookStylesPart stylesPart = _document.WorkbookPart.WorkbookStylesPart;
            Stylesheet StyleSheet = stylesPart.Stylesheet;

            // шрифты
            Fonts Fonts = StyleSheet.Elements<Fonts>().First();
            if (ExcelStyle.CellFont.Id is null)
            {
                ExcelStyle.CellFont.Append();
                Fonts.Append(ExcelStyle.CellFont.Font);
                ExcelStyle.CellFont.Id = Fonts.Count++;
            }

            // заливка
            Fills Fills = StyleSheet.Elements<Fills>().First();
            if (ExcelStyle.CellFill.Id is null)
            {
                ExcelStyle.CellFill.Append();
                Fills.Append(ExcelStyle.CellFill.Fill);
                ExcelStyle.CellFill.Id = Fills.Count++;
            }

            // границы
            Borders Borders = StyleSheet.Elements<Borders>().First();
            if (ExcelStyle.CellBorder.Id is null)
            {
                ExcelStyle.CellBorder.Append();
                Borders.Append(ExcelStyle.CellBorder.Border);
                ExcelStyle.CellBorder.Id = Borders.Count++;
            }

            CellFormats CellFormats = StyleSheet.Elements<CellFormats>().First();
            if (ExcelStyle.Id is null)
            {
                CellFormat CellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = ExcelStyle.CellFont.Id, FillId = ExcelStyle.CellFill.Id, BorderId = ExcelStyle.CellBorder.Id, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
                ExcelStyle.CellAlignment.Append();
                CellFormat.Append(ExcelStyle.CellAlignment.Alignment);
                CellFormats.Append(CellFormat);

                ExcelStyle.Id = CellFormats.Count++;
            }

            /*
            var newNumId = targetStyleSheet.NumberingFormats.Any()
                                 ? targetStyleSheet.NumberingFormats.Elements<NumberingFormat>()
                                     .Max(a => a.NumberFormatId).Value + 1
                                 : 1;
            */
            return ExcelStyle.Id;
        }
    }
}
