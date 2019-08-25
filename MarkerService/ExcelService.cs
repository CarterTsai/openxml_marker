using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace openxml.Services
{
    public class ExcelService: iMarkerService
    {
        public bool Create() 
        {
            var memoryStream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet() { 
                    Id = workbookPart.GetIdOfPart(worksheetPart), 
                    SheetId = 1, 
                    Name = "Sheet 1" 
                });

                // 要從 MemoryStream 匯出，必須先儲存 Workbook，並關閉 SpreadsheetDocument 物件
                workbookPart.Workbook.Save();
                document.Close();

                using (var fileStream = new FileStream("test.xlsx", FileMode.Create))
                {
                    memoryStream.WriteTo(fileStream);
                }
            }
            return true;
        }
    }
}