using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PlannerOpenXML.Services;

public class PlannerStyleService
{
    public Stylesheet GenerateStylesheet()
    {
        var workbookstylesheet = new Stylesheet();

        var font0 = new Font
        {
            FontSize = new FontSize() { Val = 17 }
        };

        var font1 = new Font
        {
            Bold = new Bold(),
            FontSize = new FontSize() { Val = 20 }
        };
        
        var font2 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "555555" } },
        };

        var font3 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "FF0000" } }
        };

        var font4 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "00008B" } },
            FontSize = new FontSize() { Val = 12 }
        };

        var font5 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "006400" } },
            FontSize = new FontSize() { Val = 12 }
        };

        var font6 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "C71585" } },
            FontSize = new FontSize() { Val = 12 }
        };

        var fonts = new Fonts(font0, font1, font2, font3, font4, font5, font6);

        Fill fill0 = new Fill();

        Fill fill1 = new Fill();
        PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFFFF" };
        patternFill1.Append(foregroundColor1);
        fill1.Append(patternFill1);

        Fill fill2 = new Fill();
        PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "DCDCDC" };
        patternFill2.Append(foregroundColor2);
        fill2.Append(patternFill2);

        Fill fill3 = new Fill();
        PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "F68072" };
        patternFill3.Append(foregroundColor3);
        fill3.Append(patternFill3);

        Fill fill4 = new Fill();
        PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "87CEFA" };
        patternFill4.Append(foregroundColor4);
        fill4.Append(patternFill4);

        Fill fill5 = new Fill();
        PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "9bff66" };
        patternFill5.Append(foregroundColor5);
        fill5.Append(patternFill5);

        Fill fill6 = new Fill();
        PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor6 = new ForegroundColor() { Rgb = "FFC0CB" };
        patternFill6.Append(foregroundColor6);
        fill6.Append(patternFill6);

        Fills fills = new Fills(fill0, fill1, fill2, fill3, fill4, fill5, fill6);

        Border border0 = new(new LeftBorder(),
                                    new RightBorder(),
                                    new TopBorder(),
                                    new BottomBorder());

        Border border1 = new Border(new LeftBorder() { Style = BorderStyleValues.Thick },
                                    new RightBorder() { Style = BorderStyleValues.Thick },
                                    new TopBorder() { Style = BorderStyleValues.Thick },
                                    new BottomBorder() { Style = BorderStyleValues.Thick });

        Border border2 = new Border(new LeftBorder() { Style = BorderStyleValues.Thick },
                                    new RightBorder() { Style = BorderStyleValues.Thick },
                                    new TopBorder() { Style = BorderStyleValues.Thin },
                                    new BottomBorder() { Style = BorderStyleValues.Thin });

        Border border3 = new Border(new LeftBorder() { Style = BorderStyleValues.Thick },
                                    new RightBorder() { Style = BorderStyleValues.Thick },
                                    new TopBorder() { Style = BorderStyleValues.Thin },
                                    new BottomBorder() { Style = BorderStyleValues.Thick });

        Borders borders = new Borders();
        borders.Append(border0);
        borders.Append(border1);
        borders.Append(border2);
        borders.Append(border3);

        CellFormat defaultStyle = new CellFormat()
        {
            FormatId = 0,
            FontId = 0,
            FillId = 0,
            BorderId = 0
        };

        Alignment alignment = new Alignment()
        {
            Horizontal = HorizontalAlignmentValues.Center,
            Vertical = VerticalAlignmentValues.Center
        };

        CellFormat nameOfMonthStyle = new CellFormat(alignment)
        {
            FontId = 1,
            BorderId = 1
        };

        CellFormat borderStyle = new CellFormat()
        {
            BorderId = 2
        };

        CellFormat saturdayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 2,
            FillId = 2,
        };

        CellFormat sundayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 3,
            FillId = 3,
        };

        CellFormat germanHolidayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 4,
            FillId = 4,
        };

        CellFormat hungarianHolidayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 5,
            FillId = 5
        };

        CellFormat germanAndHungarianHolidayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 6,
            FillId = 6
        };

        CellFormat lastDayOfMonthStyle = new CellFormat()
        {
            BorderId = 3
        };

        CellFormat lastDayOfMonthAndSaturdayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 2,
            FillId = 2
        };

        CellFormat lastDayOfMonthAndSundayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 3,
            FillId = 3
        };

        CellFormat lastDayOfMonthAndGermanHolidayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 4,
            FillId = 4
        };

        CellFormat lastDayOfMonthAndHungarianHolidayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 5,
            FillId = 5
        };

        CellFormat lastDayOfMonthAndGermanAndHungarianHolidayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 6,
            FillId = 6
        };

        CellFormats cellformats = new CellFormats();
        cellformats.Append(defaultStyle);
        cellformats.Append(nameOfMonthStyle);
        cellformats.Append(borderStyle);
        cellformats.Append(saturdayStyle);
        cellformats.Append(sundayStyle);
        cellformats.Append(germanHolidayStyle);
        cellformats.Append(hungarianHolidayStyle);
        cellformats.Append(germanAndHungarianHolidayStyle);
        cellformats.Append(lastDayOfMonthStyle);
        cellformats.Append(lastDayOfMonthAndSaturdayStyle);
        cellformats.Append(lastDayOfMonthAndSundayStyle);
        cellformats.Append(lastDayOfMonthAndGermanHolidayStyle);
        cellformats.Append(lastDayOfMonthAndHungarianHolidayStyle);
        cellformats.Append(lastDayOfMonthAndGermanAndHungarianHolidayStyle);

        workbookstylesheet.Append(fonts);
        workbookstylesheet.Append(fills);
        workbookstylesheet.Append(borders);
        workbookstylesheet.Append(cellformats);
        
        return workbookstylesheet;
    }
}

