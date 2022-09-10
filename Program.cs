using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;

namespace PhraseFromMSWordToExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var wordFilePath = "C:\\1\\MyFileWithPhrase.docx";
            var excelFilePath = "C:\\1\\ResultFileWithPhrase.xlsx";
            var columnName = "B";
            uint rowNumber = 2;

            Console.WriteLine("Копируем фразу...");

            var phrase = ReadPhraseFromFileMSWord(wordFilePath);
            Console.WriteLine(phrase);

            var file = OpenOrCreateExcelFile(excelFilePath);
            WritePhraseInExcelFile(file, phrase, columnName, rowNumber);
            CloseExcelFile(file);
            Console.WriteLine("Фраза успешно скопирована из Word в Excel!");
            Console.ReadLine();
        }

        public static string ReadPhraseFromFileMSWord(string path)
        {
            var document = WordprocessingDocument.Open(path, true);
            var body = document.MainDocumentPart.Document.Body;
            var phrase = body.InnerText;
            document.Close();
            return phrase;
        }

        public static SpreadsheetDocument OpenOrCreateExcelFile(string path)
        {
            SpreadsheetDocument document;
            if (File.Exists(path))
                document = SpreadsheetDocument.Open(path, true);
            else
                document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            return document;
        }

        public static void WritePhraseInExcelFile(SpreadsheetDocument file, string phrase, string columnName, uint rowNumber)
        {
            SharedStringTablePart shareStringPart;
            if (file.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                shareStringPart = file.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = file.WorkbookPart.AddNewPart<SharedStringTablePart>();
            var index = InsertSharedStringItem(phrase, shareStringPart);
            var worksheetPart = GetFirstOrCreateNewWorksheet(file.WorkbookPart);
            var cell = InsertCellInWorksheet(columnName, rowNumber, worksheetPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            worksheetPart.Worksheet.Save();
            file.Save();
        }

        public static int InsertSharedStringItem(string phrase, SharedStringTablePart shareStringPart)
        {
            if (shareStringPart.SharedStringTable == null)
                shareStringPart.SharedStringTable = new SharedStringTable();

            var elements = shareStringPart.SharedStringTable.Elements<SharedStringItem>();
            var element = elements.FirstOrDefault(x => x.InnerText == phrase);
            if (element != null)
                return elements.ToList().IndexOf(element);

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(phrase)));
            shareStringPart.SharedStringTable.Save();

            return elements.Count() + 1;
        }

        public static WorksheetPart GetFirstOrCreateNewWorksheet(WorkbookPart workbookPart)
        {
            if (workbookPart.GetPartsOfType<WorksheetPart>().Any())
                return workbookPart.GetPartsOfType<WorksheetPart>().First();

            var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            var sheetName = "Лист " + shepetId;

            var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }


        public static Cell InsertCellInWorksheet(string columnName, uint rowNumber, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = columnName + rowNumber.ToString();

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowNumber).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowNumber).First();
            }
            else
            {
                row = new Row() { RowIndex = rowNumber };
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell refCell = null;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                var newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }

        }

        public static void CloseExcelFile(SpreadsheetDocument file)
        {
            file.Close();
        }

    }
}
