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

       //this method is used to retrieve the part of the spread sheet we will write data to
        public static SheetData GetSheetDataPart(SpreadsheetDocument package)
        {
            var workBookPart = package.WorkbookPart;
            WorksheetPart workSheetPart = workBookPart.WorksheetParts.First();
            SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();
            return sheetData;
        }

        //the rest of the methods help us create diffetent types of cells
        public static Cell CreateInlineCell(string reference, string value)
        {
            Cell cell = new Cell { DataType = CellValues.InlineString, CellReference = reference };
            Text t = new Text { Text = value };
            InlineString inlineString = new InlineString();
            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);
            return cell;
        }

        public static Cell CreateDateCell(string reference, DateTime value)
        {
            Cell dateCell = new Cell() { CellReference = reference, StyleIndex = (UInt32Value)1U };
            CellValue cellValueDueDate = new CellValue();
            cellValueDueDate.Text = value.ToOADate().ToString();
            dateCell.Append(cellValueDueDate);
            return dateCell;
        }

        public static Cell CreateFormulaCell(string reference, string formular)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = reference,
                CellFormula = new CellFormula(formular)
            };
            return cell;
        }
    }
}
