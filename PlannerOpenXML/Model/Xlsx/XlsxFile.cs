using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace PlannerOpenXML.Model.Xlsx;

public sealed class XlsxFile : IDisposable
{
    #region fields
    private readonly SpreadsheetDocument m_SpreadSheet;
    #endregion fields

    #region properties
    internal SharedStringCache SharedStringCache { get; }

    public Sheets Sheets { get; private set; } = [];
    public bool Failed { get; }
    #endregion properties

    #region constructors
    private XlsxFile(string path, bool create, bool isEditable = false)
    {
        ArgumentNullException.ThrowIfNull(path);

        if (!create)
        {
            m_SpreadSheet = SpreadsheetDocument.Open(path, isEditable);

            var workbookPart = m_SpreadSheet.WorkbookPart ?? throw new NotSupportedException();

            if (isEditable)
            {
                workbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            }

            // string table
            var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringCache = new(sharedStringTablePart.SharedStringTable);

            foreach (DocumentFormat.OpenXml.Spreadsheet.Sheet sheet in workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
            {
                var xlsxSheet = new Sheet(this, workbookPart, sheet);
                Sheets.Add(xlsxSheet);
            }
        }
        else
        {
            try
            {
                m_SpreadSheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook, false);
            }
            catch (IOException)
            {
                m_SpreadSheet = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
                Failed = true;
                SharedStringCache = new();
                return;
            }
            m_SpreadSheet.AddWorkbookPart();
            if (m_SpreadSheet.WorkbookPart is null)
                throw new NotSupportedException();

            m_SpreadSheet.WorkbookPart.Workbook = new Workbook
            {
                Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets()
            };
            m_SpreadSheet.WorkbookPart.Workbook.Save();

            var sharedStringTablePart = m_SpreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
            sharedStringTablePart.SharedStringTable.Save();
            SharedStringCache = new(sharedStringTablePart.SharedStringTable);

            var workbookStylesPart = m_SpreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet.Save();
        }

    }
    #endregion constructors

    #region methods
    public static XlsxFile? Open(string path)
    {
        var result = new XlsxFile(path, false, true);
        if (result.Failed)
        {
            return null;
        }

        return result;
    }

    public void Dispose()
    {
        m_SpreadSheet?.Dispose();
    }
    #endregion methods
}
