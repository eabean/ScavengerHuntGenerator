﻿using System;
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
        private GameDetailsRepository _repo;
        public GameDetailsExporter(List<Game> games, string exportFolder, string resourcePath, GameDetailsRepository repo)
        {
            _games = games;
            _exportFolder = exportFolder;
            _resourcePath = resourcePath;
            _repo = repo;
        }

        public void ExportGameLegend()
        {
            var outputFileName = $"GameLegend.xlsx";
            var outputFilePath = Path.Combine(_exportFolder, outputFileName);
            Directory.CreateDirectory(_exportFolder);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Legend");
                int startRow = 2;
                // Clues
                for (int i = startRow;i <= Game.NUM_OF_CLUES +startRow; i++)
                {
                    worksheet.Cells[1, i].Value = $"C{i-1}";                  
                }

                // Games

                for (int i = startRow, g= 0; i < (_games.Count)*2; i += 2, g++)
                {
                    worksheet.Cells[i, 1].Value = _games[g].gameId + " Loc"; // Clue Locations
                    worksheet.Cells[i+1, 1].Value = _games[g].gameId +  " Ans"; // Clue Answer
                }


                for (int row = startRow, gIndex =0; row < (_games.Count) * 2; row +=2, gIndex++)
                {
                    var currentGame = _games[gIndex];
                    for (int col = 2, g = 0; col < Game.NUM_OF_CLUES + startRow; col++, g++)
                    {
                        var correctLocation = currentGame.selectedLocations[g];
                        var correctAnswer = ParseCorrectAnswer(currentGame.selectedQuestions, g);
                        worksheet.Cells[row, col+1].Value = $"{correctLocation.locId}"; 
                        worksheet.Cells[row+1, col].Value = $"{correctAnswer}"; 
                    }
                }

                var correctLocations = _repo.ParseLocations();
                var locStartCol = 15;
                // Headers
                worksheet.Cells[1, locStartCol].Value = "LocId";
                worksheet.Cells[1, locStartCol+1].Value = "decodedDescription";
                worksheet.Cells[1, locStartCol+2].Value = "clueDescription";

                for (int row = 2, i =0; row < correctLocations.Count; row++, i++)
                {
                    worksheet.Cells[row, locStartCol].Value = $"{correctLocations[i].locId}";
                    worksheet.Cells[row, locStartCol+1].Value = $"{correctLocations[i].decodedDescription}";
                    worksheet.Cells[row, locStartCol+2].Value = $"{correctLocations[i].clueDescription}";
                }


                package.SaveAs(new FileInfo(outputFilePath));
            }

            Console.WriteLine("Excel file exported successfully.");
        }

        public string ParseCorrectAnswer(List<Question> gameQuestions, int index)
        {
            var currentQuestion = gameQuestions[index];
            string correctLetter =  currentQuestion.qAnswers
                .FirstOrDefault(a => a.isCorrect).anText[0].ToString();

            return correctLetter;
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

                    if (tableCells.Length <= Game.NUM_OF_CLUES) throw new Exception("Not enough table cells to print out all clues and final instruction.");

                    var questions = game.selectedQuestions;

                    for (int i=0; i<= questions.Count; i++) 
                    {
                        if(i < questions.Count)
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
                            Run questionRun = new Run(new Text($"{game.gameId}{i + 1}. {question.qText}"));
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

                        } else
                        {
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
                            Run finalInstructionRun = new Run(new Text($"You've finished! Please return to the 27th floor kitchen" +
                                $" and submit your envelopes to secure your placing."));

                            RunProperties runProperties = new RunProperties();
                            FontSize fontSize = new FontSize() { Val = "24" };
                            runProperties.Append(fontSize);
                            finalInstructionRun.RunProperties = runProperties;
                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(new Break());
                            paragraph.Append(finalInstructionRun);


                            ParagraphProperties paragraphProperties = new ParagraphProperties();
                            Justification justification = new Justification() { Val = JustificationValues.Left };
                            paragraphProperties.Append(justification);
                            paragraph.Append(paragraphProperties);

                            cell.Append(paragraph);

                        }
                       

                     }
                    mainPart.Document.Save();
                    Console.WriteLine("Games clues processed successfully.");
                }

             }

        }
     }

            
}

