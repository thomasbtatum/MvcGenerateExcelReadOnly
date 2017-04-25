using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;

namespace WebApplication1.Controllers
{
    public class ExcelReportGenerator
    {
        internal void CreateExcelDocNew(MemoryStream mem)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook, true))
            {
                var relationshipId = "rId1";

                //build Workbook Part
                var workbookPart = document.AddWorkbookPart();
                var workbook = new Workbook();
                var sheets = new Sheets();
                var sheet1 = new Sheet() { Name = "Protected Sheet", SheetId = 1, Id = relationshipId };
                sheets.Append(sheet1);
                workbook.Append(sheets);
                workbookPart.Workbook = workbook;

                //build Worksheet Part
                var workSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshipId);
                var workSheet = new Worksheet();
                workSheet.Append(new SheetData());
                workSheet.Append(new SheetProtection() { Sheet = true, Objects = true, Scenarios = true });
                workSheetPart.Worksheet = workSheet;

                //add document properties
                document.PackageProperties.Creator = "Ctrl_Alt_Defeat";
                document.PackageProperties.ContentStatus = "Final";
                document.PackageProperties.Created = DateTime.UtcNow;

            }
        }
    }
}