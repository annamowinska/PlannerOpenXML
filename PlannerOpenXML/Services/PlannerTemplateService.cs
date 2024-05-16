using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PlannerOpenXML.Services;

public class PlannerTemplateService
{
    public Stylesheet ReadStylesheetFromExcel(string filePath)
    {
        Stylesheet stylesheet = null;

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart?.WorkbookStylesPart?.Stylesheet != null)
            {
                stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
            }
        }

        return stylesheet;
    }
}
