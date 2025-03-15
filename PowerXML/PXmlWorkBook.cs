using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PowerXML;

public class PXmlWorkBook
{
    public FileData CreateFile()
    {
        var fileData = new FileData();

        // Используем MemoryStream для создания файла в памяти
        var memoryStream = new MemoryStream();
            
        // Создаем новый документ Excel
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
        // Добавляем WorkbookPart
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Добавляем WorksheetPart
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Добавляем Sheets
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

        // Создаем новый лист
        Sheet sheet = new Sheet()
        {
            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        };
        sheets.Append(sheet);

        // Получаем SheetData
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        // Добавляем данные
        Row row = new Row();
        row.Append(CreateCell("Hello", CellValues.String));
        row.Append(CreateCell("World", CellValues.String));
        sheetData.Append(row);

        // Сохраняем изменения
        spreadsheetDocument.Save();

        fileData.Data = memoryStream.ToArray();

        return fileData;
    }

    private Cell CreateCell(string text, CellValues dataType)
    {
        return new Cell()
        {
            CellValue = new CellValue(text),
            DataType = new EnumValue<CellValues>(dataType)
        };
    }
}

public class FileData
{
    public byte[] Data { get; set; }
}