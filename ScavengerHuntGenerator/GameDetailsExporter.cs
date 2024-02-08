using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Word = Microsoft.Office.Interop.Word;


namespace ScavengerHuntGenerator
{
    public class GameDetailsExporter
    {

        private string _exportPath;
        public GameDetailsExporter(string exportPath)
        {
            _exportPath = exportPath;
        }

        public void CreateWordDocument()
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(_exportPath, WordprocessingDocumentType.Document))
            {
                // Add a new main document part
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();

                // Add a body to the document
                Body body = mainPart.Document.AppendChild(new Body());

                // Create a table with 3 rows and 3 columns
                Table table = new Table();
                for (int i = 0; i < 3; i++)
                {
                    TableRow row = new TableRow();
                    for (int j = 0; j < 3; j++)
                    {
                        TableCell cell = new TableCell();
                        cell.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));

                        // Add some text inside each cell
                        Paragraph paragraph = new Paragraph(new Run(new Text($"Cell {i + 1}-{j + 1}")));
                        cell.Append(paragraph);

                        row.Append(cell);
                    }
                    table.Append(row);
                }

                // Set the width of the table to take up the majority of the page
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
                table.AppendChild(new TableProperties(tableWidth));

                // Add the table to the document body
                body.Append(table);
            }

        }
    }
}
