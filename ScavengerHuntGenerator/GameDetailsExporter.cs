using Color = System.Drawing.Color;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using Path = System.IO.Path;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;

namespace ScavengerHuntGenerator
{
    public class GameDetailsExporter
    {
        private readonly List<Game> _games;
        private readonly string _exportFolder;
        private readonly string _resourcePath;
        private readonly GameDetailsRepository _repo;
        private readonly GameSettings _settings;

        private const int LegendLocationReferenceCol = 15;

        private static readonly List<Color> LocationColorPalette = new()
        {
            Color.FromArgb(255, 179, 186), // pink
            Color.FromArgb(255, 223, 186), // peach
            Color.FromArgb(255, 255, 186), // yellow
            Color.FromArgb(186, 255, 201), // green
            Color.FromArgb(186, 225, 255), // blue
            Color.FromArgb(218, 186, 255), // purple
            Color.FromArgb(255, 186, 240), // magenta
            Color.FromArgb(186, 255, 255), // cyan
            Color.FromArgb(255, 214, 165), // orange
            Color.FromArgb(165, 255, 214), // mint
            Color.FromArgb(200, 255, 165), // lime
            Color.FromArgb(240, 165, 255), // lavender
        };

        public GameDetailsExporter(List<Game> games, string exportFolder, string resourcePath, GameDetailsRepository repo, GameSettings settings)
        {
            _games = games;
            _exportFolder = exportFolder;
            _resourcePath = resourcePath;
            _repo = repo;
            _settings = settings;
        }

        public void ExportGameLegend()
        {
            Directory.CreateDirectory(_exportFolder);
            var outputFilePath = Path.Combine(_exportFolder, "GameLegend.xlsx");

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Legend");

            var locations = _repo.ParseLocations();
            var locationColors = locations
                .Select((loc, i) => (loc.locId, LocationColorPalette[i % LocationColorPalette.Count]))
                .ToDictionary(x => x.locId, x => x.Item2);

            WriteClueHeaders(ws);
            WriteGameRows(ws, locationColors);
            WriteLocationReference(ws, locations, locationColors);

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            package.SaveAs(new FileInfo(outputFilePath));
            Console.WriteLine("Excel file exported successfully.");
        }

        private void WriteClueHeaders(ExcelWorksheet ws)
        {
            // NumOfClues + 1 columns: C1 has no location, C(last) has no answer
            for (int clue = 1; clue <= _settings.NumOfClues + 1; clue++)
                ws.Cells[1, clue + 1].Value = $"Clue {clue}";
        }

        private void WriteGameRows(ExcelWorksheet ws, Dictionary<string, Color> locationColors)
        {
            for (int g = 0; g < _games.Count; g++)
            {
                int row = 2 + g * 2;
                var game = _games[g];

                ws.Cells[row, 1].Value = $"Game {game.gameId} Location";
                ws.Cells[row + 1, 1].Value = $"Game {game.gameId} Answer";

                // i iterates over NumOfClues+1 columns (C1 to C(N+1))
                // location[i-1] is where clue card i+1 is placed (answer[i] leads there)
                for (int i = 0; i <= _settings.NumOfClues; i++)
                {
                    int col = i + 2;
                    bool isFirstClue = i == 0;
                    bool isLastClue = i == _settings.NumOfClues;

                    if (!isFirstClue)
                    {
                        var locId = game.selectedLocations[i - 1].locId;
                        var cell = ws.Cells[row, col];
                        cell.Value = locId;
                        if (locationColors.TryGetValue(locId, out var color))
                            SetCellColor(cell, color);
                    }

                    if (!isLastClue)
                        ws.Cells[row + 1, col].Value = GetCorrectAnswerLetter(game.selectedQuestions, i);
                }
            }
        }

        private void WriteLocationReference(ExcelWorksheet ws, List<Location> locations, Dictionary<string, Color> locationColors)
        {
            int col = LegendLocationReferenceCol;

            ws.Cells[1, col].Value = "LocationId";
            ws.Cells[1, col + 1].Value = "decodedDescription";
            ws.Cells[1, col + 2].Value = "clueDescription";

            for (int i = 0; i < locations.Count; i++)
            {
                var locIdCell = ws.Cells[i + 2, col];
                locIdCell.Value = locations[i].locId;
                if (locationColors.TryGetValue(locations[i].locId, out var color))
                    SetCellColor(locIdCell, color);

                ws.Cells[i + 2, col + 1].Value = locations[i].decodedDescription;
                ws.Cells[i + 2, col + 2].Value = locations[i].clueDescription;
            }
        }

        private static void SetCellColor(ExcelRange cell, Color color)
        {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        private string GetCorrectAnswerLetter(List<Question> questions, int index)
        {
            return questions[index].qAnswers.First(a => a.isCorrect).anText[0].ToString();
        }

        public void ExportClues()
        {
            Directory.CreateDirectory(_exportFolder);
            foreach (var game in _games)
            {
                var outputFilePath = Path.Combine(_exportFolder, $"Game{game.gameId}.docx");
                File.Copy(_resourcePath, outputFilePath, true);

                using var doc = WordprocessingDocument.Open(outputFilePath, true);
                var table = doc.MainDocumentPart.Document.Body.Descendants<Table>().First();
                var cells = table.Descendants<TableCell>().ToArray();

                if (cells.Length <= _settings.NumOfClues)
                    throw new Exception("Not enough table cells to print out all clues and final instruction.");

                for (int i = 0; i < game.selectedQuestions.Count; i++)
                    WriteClueCell(cells[i], game, i);

                WriteFinalInstructionCell(cells[game.selectedQuestions.Count]);

                doc.MainDocumentPart.Document.Save();
                Console.WriteLine($"Game {game.gameId} clues exported successfully.");
            }
        }

        private void WriteClueCell(TableCell cell, Game game, int index)
        {
            var question = game.selectedQuestions[index];
            cell.RemoveAllChildren();
            cell.Append(BuildCellProperties());

            var questionRun = BuildRun($"{game.gameId}{index + 1}. {question.qText}");
            questionRun.Append(new Break(), new Break());

            var answersRun = BuildRun();
            foreach (var entry in question.answersLocationMapping)
            {
                answersRun.Append(new Text(entry.Key), new Break());
                answersRun.Append(new Text($"→ {entry.Value}"), new Break());
            }

            cell.Append(BuildParagraph(questionRun, answersRun));
        }

        private void WriteFinalInstructionCell(TableCell cell)
        {
            cell.RemoveAllChildren();
            cell.Append(BuildCellProperties());
            cell.Append(BuildParagraph(BuildRun(_settings.FinalInstruction)));
        }

        private static TableCellProperties BuildCellProperties()
        {
            int marginWidth = (int)(0.5 * 1440);
            var margin = new TableCellMargin
            {
                LeftMargin = new LeftMargin { Width = marginWidth.ToString() },
                RightMargin = new RightMargin { Width = marginWidth.ToString() }
            };
            return new TableCellProperties(margin);
        }

        private static Run BuildRun(string text = "")
        {
            var run = new Run();
            if (!string.IsNullOrEmpty(text))
                run.Append(new Text(text));
            run.RunProperties = new RunProperties(new FontSize { Val = "24" });
            return run;
        }

        private static Paragraph BuildParagraph(params Run[] runs)
        {
            var paragraph = new Paragraph();
            paragraph.Append(new Break(), new Break(), new Break());
            paragraph.Append(runs);
            paragraph.Append(new ParagraphProperties(new Justification { Val = JustificationValues.Left }));
            return paragraph;
        }
    }
}
