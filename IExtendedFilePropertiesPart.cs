using DocumentFormat.OpenXml.Packaging;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    public interface IExtendedFilePropertiesPart
    {
        /// <summary>
        /// Интерфейс для переопределения алгоритма создания ExtendedFilePropertiesPart для документа
        /// </summary>
        public void GenerateExtendedFilePropertiesPart(string SheetName, SpreadsheetDocument Document);
    }
}
