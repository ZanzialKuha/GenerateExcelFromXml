using DocumentFormat.OpenXml.Packaging;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    public interface IWorkbookStylesPart
    {
        /// <summary>
        /// Интерфейс для переопределения алгоритма создания набора стилей (WorkbookStylesPart) документа
        /// </summary>
        public void GenerateWorkbookStylesPart(SpreadsheetDocument Document);
    }
}
