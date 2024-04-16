using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
//using System.Windows.Media;
using System.Collections.Generic;
using Microsoft.Win32;
using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
//using System.Windows.Controls;

namespace Planner.Model
{
    public class PlannerGenerator
    {
        public async Task GeneratePlanner(int? year, int? firstMonth, int? numberOfMonths)
        {
            if (year == null || firstMonth == null || numberOfMonths == null)
            {
                MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Select file path to save the planner";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "Planner";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(saveFileDialog.FileName, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylePart.Stylesheet = GenerateStylesheet();
                        stylePart.Stylesheet.Save();

                        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Planner" };
                        sheets.Append(sheet);

                        CultureInfo culture = new CultureInfo("de-DE");
                        //CultureInfo culture = CultureInfo.CurrentCulture;

                        for (int i = 0; i < numberOfMonths; i++)
                        {
                            int currentYear = year.Value;
                            int currentMonth = firstMonth.Value + i;

                            while (currentMonth > 12)
                            {
                                currentYear++;
                                currentMonth -= 12;
                            }

                            DateTime monthDate = new DateTime(currentYear, currentMonth, 1);
                            string monthName = monthDate.ToString("MMMM yyyy", culture);
                            
                            uint styleIndex = 0;
                            Cell monthYearCell = new Cell(new CellValue($"{monthName}"))
                            {
                                DataType = CellValues.String,
                                StyleIndex = styleIndex
                            };
                             
                            AppendCellToWorksheet(spreadsheetDocument, worksheetPart, monthYearCell, 1, (uint)(i + 1));

                            int currentRow = 2;

                            DateTime currentDate = monthDate;
                            while (currentDate.Month == monthDate.Month)
                            {
                                string dayOfWeek = currentDate.ToString("ddd", culture);
                                string cellValue = $"{currentDate.Day} {dayOfWeek}";
                                Cell dateCell = new Cell(new CellValue(cellValue))
                                {
                                    DataType = CellValues.String,
                                };
                                AppendCellToWorksheet(spreadsheetDocument, worksheetPart, dateCell, (uint)currentRow, (uint)(i + 1));
                                
                                currentDate = currentDate.AddDays(1);
                                currentRow++;
                            }

                            SetColumnWidth(worksheetPart, (uint)(i + 1), 30);
                        }

                        workbookPart.Workbook.Save();
                    }

                    MessageBox.Show($"Planner has been generated and saved as: {saveFileDialog.FileName}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while generating the planner: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        #region AppendCellToWorksheet
        private void AppendCellToWorksheet(SpreadsheetDocument spreadsheetDocument, WorksheetPart worksheetPart, Cell cell, uint rowIndex, uint columnIndex)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            else
            {
                Cell existingCell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == $"{GetColumnName(columnIndex)}{rowIndex}");

                if (existingCell != null)
                {
                    row.RemoveChild(existingCell);
                }
            }
            cell.CellReference = $"{GetColumnName(columnIndex)}{rowIndex}";
            row.Append(cell);
        }
        #endregion

        private string GetColumnName(uint columnIndex)
        {
            uint dividend = columnIndex;
            string columnName = String.Empty;
            uint modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (uint)((dividend - modifier) / 26);
            }

            return columnName;
        }

        private static void SetColumnWidth(WorksheetPart worksheetPart, uint columnIndex, double width)
        {
            DocumentFormat.OpenXml.Spreadsheet.Columns columns = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (columns == null)
            {
                columns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                worksheetPart.Worksheet.InsertAt(columns, 0);
            }

            DocumentFormat.OpenXml.Spreadsheet.Column column = new DocumentFormat.OpenXml.Spreadsheet.Column()
            {
                Min = columnIndex,
                Max = columnIndex,
                Width = width,
                CustomWidth = true
            };

            columns.Append(column);
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = new Stylesheet();

            Borders borders = new Borders(
            new Border( // index 1 black border
                new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin })
        );
            DocumentFormat.OpenXml.Spreadsheet.Font boldFont = new DocumentFormat.OpenXml.Spreadsheet.Font(new DocumentFormat.OpenXml.Spreadsheet.Bold());
            
            CellFormats cellFormats = new CellFormats();
              CellFormat boldCellFormat = new CellFormat()
              {
                FontId = 0,
                BorderId = 1
              };
              cellFormats.Append(boldCellFormat);
              styleSheet.CellFormats = cellFormats;

            return styleSheet;
        }
    }
}
