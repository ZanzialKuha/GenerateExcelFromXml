using DocumentFormat.OpenXml.Packaging;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    public interface IWorkbookPart
    {
        /// <summary>
        /// Интерфейс для переопределения алгоритма создания WorkbookPart и WorksheetPart для документа
        /// </summary>
        public void GenerateWorkbookPart(string SheetName, SpreadsheetDocument Document);
    }
}
