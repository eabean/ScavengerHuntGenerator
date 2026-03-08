using OfficeOpenXml;

namespace ScavengerHuntGenerator
{
    public class GameDetailsRepository
    {
        private readonly string _pathToExcel;

        public GameDetailsRepository(string pathToExcel)
        {
            if (!File.Exists(pathToExcel))
                throw new FileNotFoundException($"Database file not found: {pathToExcel}");
            _pathToExcel = pathToExcel;
        }

        public List<Question> ParseQuestions()
        {
            var questions = new List<Question>();
            using var package = new ExcelPackage(new FileInfo(_pathToExcel));
            var ws = package.Workbook.Worksheets[SheetIndex.Questions];
            int rowCount = GetLastNonEmptyRow(ws);

            for (int row = Col.StartRow; row <= rowCount; row++)
            {
                questions.Add(new Question
                {
                    qId = int.Parse(ws.Cells[row, Col.QId].Value?.ToString()),
                    qText = ws.Cells[row, Col.QText].Value?.ToString(),
                    qAnswers = ParseAnswers(ws.Cells[row, Col.QAnswers].Value?.ToString())
                });
            }

            foreach (var q in questions)
                Console.WriteLine($"Id: {q.qId}, Text: {q.qText}, Answers: [{string.Join(", ", q.qAnswers)}]");

            return questions;
        }

        public List<Answer> ParseAnswers(string answerBlock)
        {
            if (string.IsNullOrEmpty(answerBlock))
                throw new Exception("Answers are empty for a question");

            return answerBlock
                .Split(new[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => new Answer
                {
                    isCorrect = item.Contains('*'),
                    anText = item.TrimEnd('*')
                })
                .ToList();
        }

        public List<Location> ParseLocations()
        {
            var locations = new List<Location>();
            using var package = new ExcelPackage(new FileInfo(_pathToExcel));
            var ws = package.Workbook.Worksheets[SheetIndex.Locations];
            int rowCount = GetLastNonEmptyRow(ws);

            for (int row = Col.StartRow; row <= rowCount; row++)
            {
                locations.Add(new Location
                {
                    locId = ws.Cells[row, Col.LId].Value?.ToString(),
                    decodedDescription = ws.Cells[row, Col.LDecodedDescription].Value?.ToString(),
                    clueDescription = ws.Cells[row, Col.LClueDescription].Value?.ToString()
                });
            }

            foreach (var loc in locations)
                Console.WriteLine($"Id: {loc.locId}, decodedDescription: {loc.decodedDescription}, clueDescription: {loc.clueDescription}");

            return locations;
        }

        public List<Location> ParseFakeLocations()
        {
            var fakeLocations = new List<Location>();
            using var package = new ExcelPackage(new FileInfo(_pathToExcel));
            var ws = package.Workbook.Worksheets[SheetIndex.FakeLocations];
            int rowCount = GetLastNonEmptyRow(ws);

            for (int row = Col.StartRow; row <= rowCount; row++)
            {
                fakeLocations.Add(new Location
                {
                    locId = ws.Cells[row, Col.FId].Value?.ToString(),
                    decodedDescription = "",
                    clueDescription = ws.Cells[row, Col.FClueDescription].Value?.ToString()
                });
            }

            foreach (var loc in fakeLocations)
                Console.WriteLine($"Id: {loc.locId}, clueDescription: {loc.clueDescription}");

            return fakeLocations;
        }

        private int GetLastNonEmptyRow(ExcelWorksheet ws)
        {
            int row = ws.Dimension.End.Row;
            while (string.IsNullOrEmpty(ws.Cells[row, 1].Value?.ToString()))
                row--;
            return row;
        }

        private static class SheetIndex
        {
            public const int Questions = 0;
            public const int Locations = 1;
            public const int FakeLocations = 2;
        }

        private static class Col
        {
            public const int StartRow = 2;
            public const int QId = 1;
            public const int QText = 2;
            public const int QAnswers = 3;
            public const int LId = 1;
            public const int LDecodedDescription = 2;
            public const int LClueDescription = 3;
            public const int FId = 1;
            public const int FClueDescription = 2;
        }
    }
}
