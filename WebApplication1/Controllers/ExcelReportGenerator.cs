using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

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
                var sheet1 = new Sheet() { Name = "Walter's Sheet", SheetId = 1, Id = relationshipId };
                sheets.Append(sheet1);
                workbook.Append(sheets);
                workbookPart.Workbook = workbook;

                //build Worksheet Part
                var workSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshipId);
                var workSheet = new Worksheet();
                workSheet.Append(new SheetData());
                workSheetPart.Worksheet = workSheet;

                //add document properties
                document.PackageProperties.Creator = "Ctrl_Alt_Defeat";
                document.PackageProperties.ContentStatus = "Final";
                document.PackageProperties.Created = DateTime.UtcNow;



                CustomFilePropertiesPart customFilePropertiesPart1 = document.AddNewPart<CustomFilePropertiesPart>("rId4");
                GenerateCustomFilePropertiesPart1Content(customFilePropertiesPart1);

            }
        }



        private void GenerateCustomFilePropertiesPart1Content(CustomFilePropertiesPart customFilePropertiesPart1)
        {
            Op.Properties properties2 = new Op.Properties();
            properties2.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            Op.CustomDocumentProperty customDocumentProperty1 = new Op.CustomDocumentProperty() { FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", PropertyId = 2, Name = "_MarkAsFinal" };
            Vt.VTBool vTBool1 = new Vt.VTBool();
            vTBool1.Text = "true";

            customDocumentProperty1.Append(vTBool1);

            properties2.Append(customDocumentProperty1);

            customFilePropertiesPart1.Properties = properties2;
        }

    }
}