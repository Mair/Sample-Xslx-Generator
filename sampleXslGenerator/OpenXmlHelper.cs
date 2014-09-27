using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sampleXslGenerator
{
    public class OpenXmlHelper
    {
        public static SheetData GetSheetDataPart(SpreadsheetDocument package)
        {
            var workBookPart = package.WorkbookPart;
            WorksheetPart workSheetPart = workBookPart.WorksheetParts.First();
            SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();
            return sheetData;
        }

        public static Cell CreateInlineCell(string refrence, string value)
        {
            Cell cell = new Cell { DataType = CellValues.InlineString, CellReference = refrence };
            Text t = new Text { Text = value };
            InlineString inlineString = new InlineString();
            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);
            return cell;
        }

        public static Cell CreateDateCell(string refrence, DateTime value)
        {
            Cell dateCell = new Cell() { CellReference = refrence, StyleIndex = (UInt32Value)1U };
            CellValue cellValueDueDate = new CellValue();
            cellValueDueDate.Text = value.ToOADate().ToString();
            dateCell.Append(cellValueDueDate);
            return dateCell;
        }

        public static Cell CreateFormulaCell(string refrence, string formular)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = refrence,
                CellFormula = new CellFormula(formular)
            };
            return cell;
        }
    }
}
