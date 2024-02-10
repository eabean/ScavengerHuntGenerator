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
