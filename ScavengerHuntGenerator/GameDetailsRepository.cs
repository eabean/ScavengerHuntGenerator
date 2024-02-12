using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    var q = new Question();

                    q.qId = int.Parse(worksheet.Cells[row, 1].Value?.ToString());
                    q.qText = worksheet.Cells[row, 2].Value?.ToString();
                    var answerBlock = worksheet.Cells[row, 3].Value?.ToString();
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

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {

                    var loc = new Location();

                    loc.locId = worksheet.Cells[row, 1].Value?.ToString();
                    loc.decodedDescription = worksheet.Cells[row, 2].Value?.ToString();
                    loc.clueDescription = worksheet.Cells[row, 3].Value?.ToString();

                    locations.Add(loc);
                }
            }

            foreach (var loc in locations)
            {
                Console.WriteLine($"Id: {loc.locId}, decodedDescription: {loc.decodedDescription}, clueDescription: {loc.clueDescription}");
            }

            return locations;
        }

    }
}
