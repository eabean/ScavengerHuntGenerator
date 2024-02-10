using System;
using System.Drawing.Printing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;


namespace ScavengerHuntGenerator
{
    public class GameDetailsExporter
    {

        private string _exportFolder;
        private string _resourcePath;
        public GameDetailsExporter(string exportFolder, string resourcePath)
        {
            _exportFolder = exportFolder;
            _resourcePath = resourcePath;

        }

        public void ExportGameLegend()
        {

        }

        public void ExportClues()
        {
            var outputFileName = "GameA.docx";

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

                            paragraph.Append(question,answers);

                            //Run a1 = new Run(new Text("First answer"));
                            //a1.Append(new Break());
                            //Run a2 = new Run(new Text("Second answer"));
                            //a2.Append(new Break());
                            //Run a3 = new Run(new Text("Third answer"));
                            //a3.Append(new Break());
                            //Run a4 = new Run(new Text("Fourth answer"));

                            //paragraph.Append(question,a1,a2,a3,a4);

                            // Center align the paragraph within the cell
                            ParagraphProperties paragraphProperties = new ParagraphProperties();
                            Justification justification = new Justification() { Val = JustificationValues.Center };
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
