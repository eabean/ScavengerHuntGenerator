using System;
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

          
            var outputFolderPath = _exportFolder + "Games";
            var outputFileName = "GameA.docx";

            var outputFilePath = Path.Combine(outputFolderPath, outputFileName);

            Directory.CreateDirectory(outputFolderPath);

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
                            Run run = new Run(new Text("How many holes are in a straw?"));
                            run.Append(new Break()); // Add a line break
                            run.Append(new Break()); // Add a line break
                            run.Append(new Text("First answer"));
                            run.Append(new Text("Second answer"));
                            run.Append(new Text("Third answer"));
                            run.Append(new Text("Fourth answer"));


                            // Center align the text
                            ParagraphProperties paragraphProperties = new ParagraphProperties();
                            Justification justification = new Justification() { Val = JustificationValues.Center };
                            paragraphProperties.Append(justification);
                            paragraph.Append(paragraphProperties);

                            paragraph.Append(run);
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
