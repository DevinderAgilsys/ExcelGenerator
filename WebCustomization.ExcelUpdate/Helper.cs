using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace WebCustomization.ExcelUpdate
{
    public static class Helper
    {
        /// <summary>
        /// Generate excel file based on the json send by the web client.
        /// </summary>
        /// <param name="roots"></param> 
        /// <returns></returns>
        public static MemoryStream generateExcel(List<Root> roots)
        {
            MemoryStream memoryStream = new MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                WorkbookStylesPart workStylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workStylePart.Stylesheet = GenerateStylesheetDefault();
                workStylePart.Stylesheet.Save();

                int index = workStylePart.Stylesheet.CellFormats.ChildElements.Count - 1;

                uint sheetId = 1;
                foreach (var item in roots)
                {
                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();
                    var sheetName = new StringValue(item.SheetName);

                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = sheetName };

                    sheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    sheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    Worksheet worksheet = new Worksheet();
                    SheetData sheetData = new SheetData();

                    Columns columns = new Columns();

                    columns.Append(new Column() { Min = 1, Max = 3, Width = 50, CustomWidth = true });
                    columns.Append(new Column() { Min = 4, Max = 4, Width = 30, CustomWidth = true });

                    worksheet.Append(columns);

                    if (!item.IsCardView)
                    {
                        if (item.Sections != null && item.Sections.Count > 0)
                        {
                            foreach (var section in item.Sections)
                            {
                                foreach (var cardDetail in section.CardDetails)
                                {

                                    Row newRow = new Row();
                                    Cell cell = new Cell();
                                    cell.DataType = CellValues.String;
                                    cell.CellValue = new CellValue(cardDetail.CardName);
                                    cell.StyleIndex = Convert.ToUInt32(3);

                                    newRow.Append(cell);
                                    sheetData.AppendChild(newRow);

                                    foreach (var field in cardDetail.Fields)
                                    {
                                        Row cardRow = new Row();
                                        Cell fieldName = new Cell();
                                        fieldName.DataType = CellValues.String;
                                        fieldName.CellValue = new CellValue(field.InputLabel);
                                        fieldName.StyleIndex = Convert.ToUInt32(4);
                                        Cell fieldValue = new Cell();
                                        fieldValue.DataType = CellValues.String;
                                        fieldValue.CellValue = new CellValue(field.ControlValue);
                                        fieldValue.StyleIndex = Convert.ToUInt32(2);

                                        cardRow.Append(fieldName);
                                        cardRow.Append(fieldValue);

                                        sheetData.Append(cardRow);
                                    }
                                }

                                for (int i = 0; i < 10; i++)
                                {
                                    Row blankRow = new Row();
                                    Cell blankCell = new Cell();
                                    blankRow.Append(blankCell);
                                    sheetData.Append(blankRow);
                                }

                            }
                        }
                    }
                    else
                    {
                        if (item.CardDetails != null && item.CardDetails.Count > 0)
                        {
                            foreach (var cardDetail in item.CardDetails)
                            {

                                Row newRow = new Row();
                                Cell cell = new Cell();
                                cell.DataType = CellValues.String;
                                cell.CellValue = new CellValue(cardDetail.CardName);
                                cell.StyleIndex = Convert.ToUInt32(3);

                                newRow.Append(cell);
                                sheetData.AppendChild(newRow);

                                foreach (var field in cardDetail.Fields)
                                {
                                    Row cardRow = new Row();
                                    Cell fieldName = new Cell();
                                    fieldName.DataType = CellValues.String;
                                    fieldName.CellValue = new CellValue(field.InputLabel);
                                    fieldName.StyleIndex = Convert.ToUInt32(4);
                                    Cell fieldValue = new Cell();
                                    fieldValue.DataType = CellValues.String;
                                    fieldValue.CellValue = new CellValue(field.ControlValue);
                                    fieldValue.StyleIndex = Convert.ToUInt32(2);

                                    cardRow.Append(fieldName);
                                    cardRow.Append(fieldValue);

                                    sheetData.Append(cardRow);
                                }

                                for (int i = 0; i < 2; i++)
                                {
                                    Row blankRow = new Row();
                                    Cell blankCell = new Cell();
                                    blankRow.Append(blankCell);
                                    sheetData.Append(blankRow);
                                }
                            }
                        }

                    }
                    worksheet.Append(sheetData);
                    worksheetPart.Worksheet = worksheet;

                    sheets.Append(sheet);
                }

                workbookPart.Workbook.Save();
                document.Close();
            }

            return memoryStream;
        }

        /// <summary>
        /// Creates style for the sheet.
        /// </summary>
        /// <returns></returns>
        private static Stylesheet GenerateStylesheetDefault()
        {
            Stylesheet excelStylesheet = new Stylesheet();
            Fonts fonts = new Fonts(
                new Font(
                    new FontSize() { Val = 11D },
                    new Color() { Theme = (UInt32Value)1U },
                     new FontName() { Val = "Calibri" },
                     new FontFamilyNumbering() { Val = 2 },
                     new FontScheme() { Val = FontSchemeValues.Minor }
                    ),
                new Font(
                    new Bold() { },
                    new FontSize() { Val = 11D },
                    new Color() { Theme = (UInt32Value)1U },
                     new FontName() { Val = "Calibri" },
                     new FontFamilyNumbering() { Val = 2 },
                     new FontScheme() { Val = FontSchemeValues.Minor }
                    )
                );


            Borders allBorders = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            allBorders.Append(border1);
            allBorders.Append(border2);


            Borders borders = new Borders(
                new Border(
                    new LeftBorder(
                        new Color() { Auto = true, Indexed = (UInt32Value)64U }
                    )
                    { Style = BorderStyleValues.Medium },
                    new RightBorder(
                        new Color() { Auto = true, Indexed = (UInt32Value)64U }
                    )
                    { Style = BorderStyleValues.Thin },
                    new TopBorder(
                        new Color() { Auto = true, Indexed = (UInt32Value)64U }
                    )
                    { Style = BorderStyleValues.Thin },
                    new BottomBorder(
                        new Color() { Auto = true, Indexed = (UInt32Value)64U }
                    )
                    { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
                );
            Fills fills = new Fills();

            fills = new Fills(
                new Fill(
                    new PatternFill() { PatternType = PatternValues.None }), //0
                new Fill(
                    new PatternFill() { PatternType = PatternValues.Gray125 }),//1
                new Fill(
                    new PatternFill(
                        new ForegroundColor { Rgb = new HexBinaryValue() { Value = "BFBFBF" } }) //2 -Grey
                    { PatternType = PatternValues.Solid }),
                new Fill(
                    new PatternFill(
                        new ForegroundColor { Rgb = new HexBinaryValue() { Value = "B4C6E7" } })// 3 -Light Blue
                    { PatternType = PatternValues.Solid })
                );

            CellStyleFormats cellStyleFormats = new CellStyleFormats(
                new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U }
                );

            CellFormats cellFormats = new CellFormats(
                new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U },
                 new CellFormat(
                     new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                 { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                 new CellFormat(
                     new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center })
                 { FontId = 0, FillId = 2, BorderId = 1, ApplyFill = true },
                new CellFormat(
                     new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                { FontId = 1, FillId = 3, BorderId = 1, ApplyFill = true },
                new CellFormat(
                     new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                { FontId = 1, FillId = 0, BorderId = 1, ApplyFill = true }
                );

            excelStylesheet.Append(fonts);
            excelStylesheet.Append(fills);
            excelStylesheet.Append(allBorders);
            excelStylesheet.Append(cellStyleFormats);
            excelStylesheet.Append(cellFormats);
            return excelStylesheet;
        }

        private static Stylesheet CreateStyle()
        {
            return new Stylesheet(
            new Fonts(
                new Font(                                                               // Index 0 - The default font.
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 11 },
                    new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                    new FontName() { Val = "Calibri" }),
                new Font(                                                               // Index 1 - The bold white font.
                    new DocumentFormat.OpenXml.Spreadsheet.Bold(),
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 11 },
                    new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = "ffffff" } },
                    new DocumentFormat.OpenXml.Spreadsheet.FontName() { Val = "Calibri" }),
                new Font(                                                               // Index 2 - The bold red font.
                    new DocumentFormat.OpenXml.Spreadsheet.Bold(),
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 11 },
                    new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = "ff0000" } },
                    new DocumentFormat.OpenXml.Spreadsheet.FontName() { Val = "Calibri" })
            ),
            new Fills(
                new Fill(                                                           // Index 0 - The default fill.
                    new PatternFill() { PatternType = PatternValues.None }),
                new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                    new PatternFill() { PatternType = PatternValues.Gray125 }),
                new Fill(                                                           // Index 2 - The blue fill.
                    new PatternFill(
                        new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "006699" } }
                    )
                    { PatternType = PatternValues.Solid }),
                 new Fill(                                                           // Index 3 - The grey fill.
                    new PatternFill(
                        new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "808080" } }
                    )
                    { PatternType = PatternValues.Solid }
                   )
            ),
            new Borders(
                new Border(                                                         // Index 0 - The default border.
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()),
                new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                    new LeftBorder(
                        new Color() { Auto = true }
                    )
                    { Style = BorderStyleValues.Medium },
                    new RightBorder(
                        new Color() { Auto = true }
                    )
                    { Style = BorderStyleValues.Thin },
                    new TopBorder(
                        new Color() { Auto = true }
                    )
                    { Style = BorderStyleValues.Thin },
                    new BottomBorder(
                        new Color() { Auto = true }
                    )
                    { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
            ),
            new CellFormats(
                new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 1, FillId = 0, BorderId = 0 },                      // Index 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead

                new CellFormat(
                                new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                                )
                { FontId = 1, FillId = 2, BorderId = 0, ApplyFont = true },   // Index 1 - Bold White Blue Fill

                new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                                )
                { FontId = 2, FillId = 2, BorderId = 0, ApplyFont = true }, // Index 2 - Bold Red Blue Fill
                new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                                )
                { FontId = 3, FillId = 3, BorderId = 0, ApplyFont = true }
            )
        );
        }
    }
}