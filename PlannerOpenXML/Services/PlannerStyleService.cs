using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PlannerOpenXML.Services;

public class PlannerStyleService
{
    public Stylesheet GenerateStylesheet()
    {
        var workbookstylesheet = new Stylesheet();

        var defaultFont = new Font
        {
            FontSize = new FontSize() { Val = 20 }
        };

        var monthFont = new Font
        {
            Bold = new Bold(),
            FontSize = new FontSize() { Val = 30 },
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } },
        };

        var dayFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "000099" } },
            FontSize = new FontSize() { Val = 20 }
        };
        
        var saturdayFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "555555" } },
            FontSize = new FontSize() { Val = 20 }
        };

        var sundayFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "FF0000" } },
            FontSize = new FontSize() { Val = 20 }
        };

        var holidayAndMilestoneFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "009900" } },
            FontSize = new FontSize() { Val = 10 }
        };

        Fonts fonts = new Fonts(defaultFont, monthFont, dayFont, saturdayFont, sundayFont, holidayAndMilestoneFont);

        Fill defaultFill = new Fill();
        PatternFill defaultPaternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor defaultForegroundColor = new ForegroundColor() { Rgb = "000000" };
        defaultPaternFill.Append(defaultForegroundColor);
        defaultFill.Append(defaultPaternFill);

        Fill sheetFill = new Fill();
        PatternFill sheetPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor sheetForegroundColor = new ForegroundColor() { Rgb = "FFFFFF" };
        sheetPatternFill.Append(sheetForegroundColor);
        sheetFill.Append(sheetPatternFill);
        
        Fill monthFill = new Fill();
        PatternFill monthPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor monthForegroundColor = new ForegroundColor() { Rgb = "000099" };
        monthPatternFill.Append(monthForegroundColor);
        monthFill.Append(monthPatternFill);

        Fill dayFill = new Fill();
        PatternFill dayPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor dayForegroundColor = new ForegroundColor() { Rgb = "99CCFF" };
        dayPatternFill.Append(dayForegroundColor);
        dayFill.Append(dayPatternFill);

        Fill saturdayFill = new Fill();
        PatternFill saturdayPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor saturdayForegroundColor = new ForegroundColor() { Rgb = "C0C0C0" };
        saturdayPatternFill.Append(saturdayForegroundColor);
        saturdayFill.Append(saturdayPatternFill);

        Fill sundayFill = new Fill();
        PatternFill sundayPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor sundayForegroundColor = new ForegroundColor() { Rgb = "F68072" };
        sundayPatternFill.Append(sundayForegroundColor);
        sundayFill.Append(sundayPatternFill);

        Fills fills = new Fills(defaultFill, sheetFill, monthFill, dayFill, saturdayFill, sundayFill);

        Border defaultBorder = new(new LeftBorder(),
                                    new RightBorder(),
                                    new TopBorder(),
                                    new BottomBorder());

        Border monthAndDayBorder = new Border(new LeftBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new RightBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new TopBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new BottomBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } });

        Border holidayAndMilestoneBorder = new Border(new LeftBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new RightBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new TopBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Rgb = "FFFFFF" } },
                                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Rgb = "000000" } });

        Borders borders = new Borders(defaultBorder, monthAndDayBorder, holidayAndMilestoneBorder);

        CellFormat defaultStyle = new CellFormat()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };

        CellFormat nameOfMonthStyle = new CellFormat()
        {
            FontId = 1,
            FillId = 2,
            BorderId = 1,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };

        CellFormat dayStyle = new CellFormat()
        {
            FontId = 2,
            FillId = 3,
            BorderId = 1,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };

        CellFormat saturdayStyle = new CellFormat()
        {
            FontId = 3,
            FillId = 4,
            BorderId = 1,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };

        CellFormat sundayStyle = new CellFormat()
        {
            FontId = 4,
            FillId = 5,
            BorderId = 1,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };

        CellFormat holidayAndMilestoneStyle = new CellFormat()
        {
            FontId = 5,
            BorderId = 2,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Bottom
            }
        };

        CellFormats cellformats = new CellFormats();
        cellformats.Append(defaultStyle);
        cellformats.Append(nameOfMonthStyle);
        cellformats.Append(dayStyle);
        cellformats.Append(saturdayStyle);
        cellformats.Append(sundayStyle);
        cellformats.Append(holidayAndMilestoneStyle);
        
        workbookstylesheet.Append(fonts);
        workbookstylesheet.Append(fills);
        workbookstylesheet.Append(borders);
        workbookstylesheet.Append(cellformats);
        
        return workbookstylesheet;
    }
}