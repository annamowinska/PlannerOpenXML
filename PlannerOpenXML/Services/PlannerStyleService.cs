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

        var firstCountryHolidayFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "009900" } },
            FontSize = new FontSize() { Val = 8 }
        };

        var secondCountryHolidayFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "660066" } },
            FontSize = new FontSize() { Val = 8 }
        };

        var milestoneFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "FF0000" } },
            FontSize = new FontSize() { Val = 15 }
        };

        var weekNumberFont = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "BCBC0D" } },
            FontSize = new FontSize() { Val = 25 }
        };

        Fonts fonts = new Fonts(defaultFont, monthFont, dayFont, saturdayFont, sundayFont, firstCountryHolidayFont, secondCountryHolidayFont, milestoneFont, weekNumberFont);

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

        Fill firstCountryHolidayFill = new Fill();
        PatternFill firstCountryHolidayPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor firstCountryHolidayForegroundColor = new ForegroundColor() { Rgb = "B2FF66" };
        firstCountryHolidayPatternFill.Append(firstCountryHolidayForegroundColor);
        firstCountryHolidayFill.Append(firstCountryHolidayPatternFill);

        Fill secondCountryHolidayFill = new Fill();
        PatternFill secondCountryHolidayPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor secondCountryHolidayForegroundColor = new ForegroundColor() { Rgb = "B583B5" };
        secondCountryHolidayPatternFill.Append(secondCountryHolidayForegroundColor);
        secondCountryHolidayFill.Append(secondCountryHolidayPatternFill);

        Fill milestoneFill = new Fill();
        PatternFill milestonePatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor milestoneForegroundColor = new ForegroundColor() { Rgb = "FFC0CB" };
        milestonePatternFill.Append(milestoneForegroundColor);
        milestoneFill.Append(milestonePatternFill);

        Fill weekNumberFill = new Fill();
        PatternFill weekNumberPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor weekNumberForegroundColor = new ForegroundColor() { Rgb = "FFFF99" };
        weekNumberPatternFill.Append(weekNumberForegroundColor);
        weekNumberFill.Append(weekNumberPatternFill);

        Fills fills = new Fills(defaultFill, sheetFill, monthFill, dayFill, saturdayFill, sundayFill, firstCountryHolidayFill, secondCountryHolidayFill, milestoneFill, weekNumberFill);

        Border defaultBorder = new(new LeftBorder(),
                                    new RightBorder(),
                                    new TopBorder(),
                                    new BottomBorder());

        Border monthAndDayBorder = new Border(new LeftBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new RightBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new TopBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new BottomBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } });

        Border holidayBorder = new Border(new LeftBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Rgb = "000000" } });

        Border milestoneBorder = new Border(new LeftBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new TopBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Rgb = "FFFFFF" } });

        Border emptyWeekNumberCellBorder = new Border(new RightBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Rgb = "000000" } });

        Border dayCellBorder = new Border(new TopBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } },
                                    new BottomBorder() { Style = BorderStyleValues.Thick, Color = new Color() { Rgb = "FFFFFF" } });

        Borders borders = new Borders(defaultBorder, monthAndDayBorder, holidayBorder, milestoneBorder, emptyWeekNumberCellBorder, dayCellBorder);

        CellFormat defaultStyle = new CellFormat() // IndexStyle = 0
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

        CellFormat nameOfMonthStyle = new CellFormat() // IndexStyle = 1
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

        CellFormat dayStyle = new CellFormat() // IndexStyle = 2
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

        CellFormat saturdayStyle = new CellFormat() // IndexStyle = 3
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

        CellFormat sundayStyle = new CellFormat() // IndexStyle = 4
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

        CellFormat firstCountryHolidayStyle = new CellFormat() // IndexStyle = 5
        {
            FontId = 5,
            BorderId = 2,
            FillId = 6,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Bottom
            }
        };

        CellFormat secondCountryHolidayStyle = new CellFormat() // IndexStyle = 6
        {
            FontId = 6,
            BorderId = 2,
            FillId = 7,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Bottom
            }
        };

        CellFormat milestoneStyle = new CellFormat() // IndexStyle = 7
        {
            FontId = 7,
            FillId = 8,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Bottom
            }
        };

        CellFormat emptyFirstCellStyle = new CellFormat() // IndexStyle = 8
        {
            BorderId = 3
        };

        CellFormat emptyFirstCellAndFirstCountryHolidayStyle = new CellFormat() // IndexStyle = 9
        {
            FillId = 6
        };

        CellFormat emptyFirstCellAndSecondCountryHolidayStyle = new CellFormat() // IndexStyle = 10
        {
            FillId = 7
        };

        CellFormat emptySecondCellStyle = new CellFormat() // IndexStyle = 11
        {
            BorderId = 2
        };

        CellFormat emptySecondCellAndMilestoneStyle = new CellFormat() // IndexStyle = 12
        {
            BorderId = 2,
            FillId = 8
        };

        CellFormat weekNumberCellStyle = new CellFormat() // IndexStyle = 13
        {
            FontId = 8,
            BorderId = 1,
            FillId = 9,
            Alignment = new Alignment()
            {
                WrapText = true,
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };

        CellFormat emptyWeekNumberCellStyle = new CellFormat() // IndexStyle = 14
        {
            BorderId = 4
        };

        CellFormat emptyWeekNumberCellAndMiletoneStyle = new CellFormat() // IndexStyle = 15
        {
            BorderId = 4,
            FillId = 8
        };

        CellFormat emptyWeekNumberCellAndFirstContryHolidayStyle = new CellFormat() // IndexStyle = 16
        {
            BorderId = 4,
            FillId = 6
        };

        CellFormat emptyWeekNumberCellAndSecondContryHolidayStyle = new CellFormat() // IndexStyle = 17
        {
            BorderId = 4,
            FillId = 7
        };

        CellFormats cellformats = new CellFormats();
        cellformats.Append(defaultStyle);
        cellformats.Append(nameOfMonthStyle);
        cellformats.Append(dayStyle);
        cellformats.Append(saturdayStyle);
        cellformats.Append(sundayStyle);
        cellformats.Append(firstCountryHolidayStyle);
        cellformats.Append(secondCountryHolidayStyle);
        cellformats.Append(milestoneStyle);
        cellformats.Append(emptyFirstCellStyle);
        cellformats.Append(emptyFirstCellAndFirstCountryHolidayStyle);
        cellformats.Append(emptyFirstCellAndSecondCountryHolidayStyle);
        cellformats.Append(emptySecondCellStyle);
        cellformats.Append(emptySecondCellAndMilestoneStyle);
        cellformats.Append(weekNumberCellStyle);
        cellformats.Append(emptyWeekNumberCellStyle);
        cellformats.Append(emptyWeekNumberCellAndMiletoneStyle);
        cellformats.Append(emptyWeekNumberCellAndFirstContryHolidayStyle);
        cellformats.Append(emptyWeekNumberCellAndSecondContryHolidayStyle);
        
        workbookstylesheet.Append(fonts);
        workbookstylesheet.Append(fills);
        workbookstylesheet.Append(borders);
        workbookstylesheet.Append(cellformats);
        
        return workbookstylesheet;
    }
}