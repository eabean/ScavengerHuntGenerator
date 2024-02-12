using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using Path = System.IO.Path;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;

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

                for(int i = 2;i < Game.NUM_OF_CLUES +2; i++)
                {
                    worksheet.Cells[1, i].Value = $"L{i-1}";                  
                }

                for (int i = 2, g= 0; i < _games.Count+2; i++, g++)
                {
                    worksheet.Cells[i, 1].Value = _games[g].gameId;
                }

                int startRow = 2;
                for (int row = startRow; row < _games.Count + startRow; row++)
                {
                    var currentGame = _games[row-2];
                    for (int col = 2; col < Game.NUM_OF_CLUES + startRow; col++)
                    {
                        var correctLocation = currentGame.selectedLocations[col - 2];
                        worksheet.Cells[row, col].Value = $"{correctLocation.locId}"; 
                    }
                }

                package.SaveAs(new FileInfo(outputFilePath));
            }

            Console.WriteLine("Excel file exported successfully.");
        }
     

        public void ExportClues()
        {
            foreach (var game in _games)
            {
                var outputFileName = $"Game{game.gameId}.docx";
                var outputFilePath = Path.Combine(_exportFolder, outputFileName);
                File.Copy(_resourcePath, outputFilePath, true);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(outputFilePath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    Table table = mainPart.Document.Body.Descendants<Table>().FirstOrDefault();
                    var tableCells = table.Descendants<TableCell>().ToArray();

                    if (tableCells.Length < Game.NUM_OF_CLUES) throw new Exception("Not enough table cells to print out all clues.");

                    var questions = game.selectedQuestions;

                    for (int i=0; i< questions.Count; i++) 
                    {

                        var question = questions[i];
                        var cell = tableCells[i];
                        cell.RemoveAllChildren();

                        int leftmarginInTwentiethsOfPoint = (int)(0.5 * 1440);
                        int rightmarginInTwentiethsOfPoint = (int)(0.5 * 1440);

                        TableCellMargin tableCellMargin = new TableCellMargin();

                        tableCellMargin.LeftMargin = new LeftMargin() { Width = leftmarginInTwentiethsOfPoint.ToString() };
                        tableCellMargin.RightMargin = new RightMargin() { Width = rightmarginInTwentiethsOfPoint.ToString() };

                        TableCellProperties cellProperties = new TableCellProperties(tableCellMargin);
                        cell.Append(cellProperties);


                        Paragraph paragraph = new Paragraph();
                        Run questionRun = new Run(new Text($"{game.gameId}{i+1}. {question.qText}"));
                        questionRun.Append(new Break());
                        questionRun.Append(new Break());
                        Run answers = new Run();

                        foreach (var alm in question.answersLocationMapping)
                        {
                            answers.Append(new Text($"{alm.Key}"));
                            answers.Append(new Break());
                            answers.Append(new Text($"→ {alm.Value}"));
                            answers.Append(new Break());
                        }
                        

                        RunProperties runProperties = new RunProperties();
                        FontSize fontSize = new FontSize() { Val = "24" };
                        runProperties.Append(fontSize);
                        RunProperties runProperties2 = new RunProperties();
                        FontSize fontSize2 = new FontSize() { Val = "24" };
                        runProperties2.Append(fontSize2);

                        questionRun.RunProperties = runProperties;
                        answers.RunProperties = runProperties2;

                        paragraph.Append(new Break());
                        paragraph.Append(new Break());
                        paragraph.Append(new Break());
                        paragraph.Append(questionRun, answers);


                        ParagraphProperties paragraphProperties = new ParagraphProperties();
                        Justification justification = new Justification() { Val = JustificationValues.Left };
                        paragraphProperties.Append(justification);
                        paragraph.Append(paragraphProperties);

                        cell.Append(paragraph);

                     }
                    mainPart.Document.Save();
                    Console.WriteLine("Games clues processed successfully.");
                }

             }

        }
     }

            
}

