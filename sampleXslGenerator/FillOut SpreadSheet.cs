using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using sampleXslGenerator;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace GeneratedCode
{
    public class PopulateSpreadSheet 
    {
         private static readonly string[] xlsxCols = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
       
        List<Person> _people;
        public PopulateSpreadSheet(List<Person> people)
        {
            _people = people;
        }

        public void Polpulate(Stream stream)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, true))
            {
                SheetData sheetData = OpenXmlHelper.GetSheetDataPart(package);
                int rowindex = 1;
                foreach (Person person in _people)
                {
                    rowindex++;
                    int coliIndex = 0;
     
                    Cell firstNameCell = OpenXmlHelper.CreateInlineCell(xlsxCols[coliIndex++] + ":" + rowindex, person.FirstName);
                    Cell lastNameCell = OpenXmlHelper.CreateInlineCell(xlsxCols[coliIndex++] + ":" + rowindex, person.LastName);
                    Cell dateOfBirthCell = OpenXmlHelper.CreateDateCell(xlsxCols[coliIndex++] + ":" + rowindex, person.DateOfBirth);
                    Cell ageCell = OpenXmlHelper.CreateFormulaCell(xlsxCols[coliIndex++] + ":" + rowindex, "Year(Now()) - Year(" + xlsxCols[coliIndex - 2] + rowindex + ")");

                    Row row = new Row();
                    row.Append(firstNameCell);
                    row.Append(lastNameCell);
                    row.Append(dateOfBirthCell);
                    row.Append(ageCell);
                    sheetData.Append(row);
                }
            }
        }
    }
}


