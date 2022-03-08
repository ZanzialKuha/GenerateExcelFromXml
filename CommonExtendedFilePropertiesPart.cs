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
    /// Вспомогательный класс для создания основной структуры документа в классе ExtendedFilePropertiesPart (см. документацию Open XML)
    /// </summary>
    public class CommonExtendedFilePropertiesPart : IExtendedFilePropertiesPart
    {
        private ExtendedFilePropertiesPart extendedFilePropertiesPart;
        private WorkbookPart workbookPart;

        /// <summary>
        /// Генерация структуры ExtendedFilePropertiesPart (см. документацию Open XML)
        /// </summary>
        public virtual void GenerateExtendedFilePropertiesPart(string SheetName, SpreadsheetDocument Document)
        {
            // создадим "заголовочную" структуру файла
            extendedFilePropertiesPart = Document.AddNewPart<ExtendedFilePropertiesPart>();

            DocumentXML.Properties properties1 = new DocumentXML.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            DocumentXML.Application application1 = new DocumentXML.Application();
            application1.Text = "Microsoft Excel";
            DocumentXML.DocumentSecurity documentSecurity1 = new DocumentXML.DocumentSecurity();
            documentSecurity1.Text = "0";
            DocumentXML.ScaleCrop scaleCrop1 = new DocumentXML.ScaleCrop();
            scaleCrop1.Text = "false";

            DocumentXML.HeadingPairs headingPairs1 = new DocumentXML.HeadingPairs();

            VariantTypes.VTVector vTVector1 = new VariantTypes.VTVector() { BaseType = VariantTypes.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            VariantTypes.Variant variant1 = new VariantTypes.Variant();
            VariantTypes.VTLPSTR vTLPSTR1 = new VariantTypes.VTLPSTR();
            vTLPSTR1.Text = "Листы";

            variant1.Append(vTLPSTR1);

            VariantTypes.Variant variant2 = new VariantTypes.Variant();
            VariantTypes.VTInt32 vTInt321 = new VariantTypes.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);
            
            DocumentXML.TitlesOfParts titlesOfParts1 = new DocumentXML.TitlesOfParts();
            VariantTypes.VTVector vTVector2 = new VariantTypes.VTVector() { BaseType = VariantTypes.VectorBaseValues.Lpstr, Size = (UInt32Value)0U };

            titlesOfParts1.Append(vTVector2);

            DocumentXML.Company company1 = new DocumentXML.Company();
            company1.Text = "";
            DocumentXML.LinksUpToDate linksUpToDate1 = new DocumentXML.LinksUpToDate();
            linksUpToDate1.Text = "false";
            DocumentXML.SharedDocument sharedDocument1 = new DocumentXML.SharedDocument();
            sharedDocument1.Text = "false";
            DocumentXML.HyperlinksChanged hyperlinksChanged1 = new DocumentXML.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            DocumentXML.ApplicationVersion applicationVersion1 = new DocumentXML.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart.Properties = properties1;

            // создадим ключевой узер WorkbookPart, содержащий структуру книги
            workbookPart = Document.AddWorkbookPart();

            Workbook workbook = new Workbook();
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "4", BuildVersion = "9302" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { FilterPrivacy = true, DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 240, YWindow = 105, WindowWidth = (UInt32Value)14805U, WindowHeight = (UInt32Value)8010U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)122211U };

            workbook.Append(fileVersion1);
            workbook.Append(workbookProperties1);
            workbook.Append(bookViews1);
            workbook.Append(sheets1);
            workbook.Append(calculationProperties1);

            workbookPart.Workbook = workbook;
        }
    }
}
