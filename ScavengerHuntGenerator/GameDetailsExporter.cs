using System;
using System.Drawing.Printing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;

namespace ScavengerHuntGenerator
{
    public class GameDetailsExporter
    {
        private List<Game> _games;
        private string _exportFolder;
        private string _resourcePath;
        public GameDetailsExporter(List<Game> games, string exportFolder, string resourcePath)
        {
            _games = games;
            _exportFolder = exportFolder;
            _resourcePath = resourcePath;   
        }

    

        public void ExportGameLegend()
        {
            var outputFileName = $"GameLegend.xlsx";
            var outputFilePath = Path.Combine(_exportFolder, outputFileName);
            Directory.CreateDirectory(_exportFolder);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Legend");

                for(int i = 2;i < Game.MAX_CLUES+2; i++)
                {
                    worksheet.Cells[1, i].Value = $"L{i-1}";                  
                }

                for (int i = 2, g= 0; i < _games.Count+1; i++, g++)
                {
                        worksheet.Cells[i, 1].Value = _games[g].gameId;
                }


                //int startRow = 2;
                //for(int row=startRow; row <= _games.Count+startRow; row++)
                //{
                //    for (int col = 1; row <= _games.Count + startRow; row++)
                //    {

                //    }
                //}
                worksheet.Cells["B2"].Value = "Hello";
                worksheet.Cells["C2"].Value = "World!";

                package.SaveAs(new FileInfo(outputFilePath));
            }

            Console.WriteLine("Excel file exported successfully.");
        }
     

        public void ExportClues(Game game)
        {
            var outputFileName = $"Game{game.gameId}.docx";

            var outputFilePath = Path.Combine(_exportFolder, outputFileName);

            Directory.CreateDirectory(_exportFolder);

            File.Copy(_resourcePath, outputFilePath, true);


            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputFilePath, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;

                foreach (Table table in mainPart.Document.Body.Descendants<Table>())
                {
                    foreach (TableRow row in table.Elements<TableRow>())
                    {
                        foreach (TableCell cell in row.Elements<TableCell>())
                        {
                            cell.RemoveAllChildren();

                            int marginInTwentiethsOfPoint = (int)(1.25 * 1440); 

                            TableCellMargin tableCellMargin = new TableCellMargin();

                            tableCellMargin.LeftMargin = new LeftMargin() { Width = marginInTwentiethsOfPoint.ToString() };
                            tableCellMargin.RightMargin = new RightMargin() { Width = marginInTwentiethsOfPoint.ToString() };

                            TableCellProperties cellProperties = new TableCellProperties(tableCellMargin);
                            cell.Append(cellProperties);


                            Paragraph paragraph = new Paragraph();
                            Run question = new Run(new Text("How many holes are in a straw?"));
                            question.Append(new Break());
                            question.Append(new Break());
                            Run answers = new Run(new Text("a1"));
                            answers.Append(new Break());
                            answers.Append(new Text("a2"));
                            answers.Append(new Break());
                            answers.Append(new Text("a3"));
                            answers.Append(new Break());
                            answers.Append(new Text("a4"));

                            RunProperties runProperties = new RunProperties();
                            FontSize fontSize = new FontSize() { Val = "24" };
                            runProperties.Append(fontSize);       
                            RunProperties runProperties2 = new RunProperties();
                            FontSize fontSize2 = new FontSize() { Val = "24" };
                            runProperties2.Append(fontSize2);

                            question.RunProperties = runProperties;
                            answers.RunProperties = runProperties2;

                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(question,answers);
                           

                            // Center align the paragraph within the cell
                            ParagraphProperties paragraphProperties = new ParagraphProperties();
                            Justification justification = new Justification() { Val = JustificationValues.Left };
                            paragraphProperties.Append(justification);
                            paragraph.Append(paragraphProperties);

                            cell.Append(paragraph);
                        }
                    }
                }
                mainPart.Document.Save();
            }

            Console.WriteLine("Word document processed successfully.");
        }
 

        }


    }
