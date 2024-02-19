using OfficeOpenXml;

namespace ScavengerHuntGenerator
{
    public class GameDetailsRepository
    {
        private string _pathToExcel;
        public GameDetailsRepository(string pathToExcel)
        {
            _pathToExcel = pathToExcel;
        }

        public List<Question> ParseQuestions()
        {
            var questions = new List<Question>();
            using (var excelPackage = new ExcelPackage(new FileInfo(_pathToExcel)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var worksheet = excelPackage.Workbook.Worksheets[0];

                int rowCount = GetLastNonEmptyRow(worksheet);
                for (int row = ExcelIndices.StartRow; row <= rowCount; row++)
                {
                    var q = new Question();

                    q.qId = int.Parse(worksheet.Cells[row, ExcelIndices.qIdCol].Value?.ToString());
                    q.qText = worksheet.Cells[row, ExcelIndices.qTextCol].Value?.ToString();
                    var answerBlock = worksheet.Cells[row, ExcelIndices.qAnswersCol].Value?.ToString();
                    q.qAnswers = ParseAnswers(answerBlock);

                    questions.Add(q);
                }
            }

            foreach (var question in questions)
            {
                Console.WriteLine($"Id: {question.qId}, Text: {question.qText}, Answer: [{string.Join(", ", question.qAnswers)}]");
            }

            return questions;
        }

        public List<Answer> ParseAnswers(string answerBlock)
        {
            var answers = new List<Answer>();
            if (!string.IsNullOrEmpty(answerBlock))
            {
                var answerArray = answerBlock.Split(new[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var answerItem in answerArray)
                {
                    var a = new Answer();
                    if(answerItem.Contains("*"))
                    {
                        a.anText = answerItem.TrimEnd('*');
                        a.isCorrect = true;
                    } else
                    {
                        a.anText = answerItem; 
                        a.isCorrect = false;
                    }
                    answers.Add(a);
                }
            } else
            {
                throw new Exception("answers are empty for a question");
            }
            return answers;
        }

        public List<Location> ParseLocations()
        {

            var locations = new List<Location>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage(new FileInfo(_pathToExcel)))
            {

                var worksheet = excelPackage.Workbook.Worksheets[1];

                int rowCount = GetLastNonEmptyRow(worksheet);
                for (int row = ExcelIndices.StartRow; row <= rowCount; row++)
                {

                    var loc = new Location();

                    loc.locId = worksheet.Cells[row, ExcelIndices.lIdCol].Value?.ToString();
                    loc.decodedDescription = worksheet.Cells[row, ExcelIndices.lddCol].Value?.ToString();
                    loc.clueDescription = worksheet.Cells[row, ExcelIndices.lcdCol].Value?.ToString();

                    locations.Add(loc);
                }
            }

            foreach (var loc in locations)
            {
                Console.WriteLine($"Id: {loc.locId}, decodedDescription: {loc.decodedDescription}, clueDescription: {loc.clueDescription}");
            }

            return locations;
        }

        public List<Location> ParseFakeLocations()
        {

            var fakeLocations = new List<Location>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage(new FileInfo(_pathToExcel)))
            {

                var worksheet = excelPackage.Workbook.Worksheets[2];

                int rowCount = GetLastNonEmptyRow(worksheet);
                for (int row = ExcelIndices.StartRow; row <= rowCount; row++)
                {

                    var loc = new Location();

                    loc.locId = worksheet.Cells[row, ExcelIndices.fIdCol].Value?.ToString();
                    loc.decodedDescription = "";
                    loc.clueDescription = worksheet.Cells[row, ExcelIndices.fcdCol].Value?.ToString();

                    fakeLocations.Add(loc);
                }
            }

            foreach (var loc in fakeLocations)
            {
                Console.WriteLine($"Id: {loc.locId}, clueDescription: {loc.clueDescription}");
            }

            return fakeLocations;
        }

        public int GetLastNonEmptyRow(ExcelWorksheet ws)
        {
            int currentLastRow = ws.Dimension.End.Row;
            while (string.IsNullOrEmpty(ws.Cells[currentLastRow, 1].Value?.ToString()))
            {
                currentLastRow--;
            }

            return currentLastRow;
        }

        public static class ExcelIndices
        {
            public const int StartRow = 2;
            public const int qIdCol = 1;
            public const int qTextCol = 2;
            public const int qAnswersCol = 3;            
            public const int lIdCol = 1;
            public const int lddCol = 2;
            public const int lcdCol = 3;
            public const int fIdCol = 1;
            public const int fcdCol = 2;

        }

    }
}
