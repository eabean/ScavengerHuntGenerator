using DocumentFormat.OpenXml.Wordprocessing;
using ScavengerHuntGenerator;
using System;


class Program
{

    static void Main(string[] args)
    {
        var executableDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var projectDirectory = Directory.GetParent(executableDirectory)?.FullName;
        var gameIds = Enumerable.Range('A', 26).Select(x => ((char)x).ToString()).ToList();

        var resourcesFolderName = "Resources";
        var resourcesFolderPath = Path.Combine(projectDirectory, resourcesFolderName);
        var dbPath = resourcesFolderPath + @"\ScavengerHuntDbShared.xlsx";

        var outputFolderName = "Output";
        string outputFolder = Path.Combine(projectDirectory, outputFolderName);
        string resourcePath = resourcesFolderPath + @"\Clues2x2.docx";

        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(dbPath);

        List<Game> gamesGenerated = new List<Game>();
        for (int i = 0; i < Game.NUM_OF_GAMES; i++)
        {
            var game = new Game(gameIds[i], gameDetailsRepository);
            game.GenerateGame();
            gamesGenerated.Add(game);   
        }

        GameDetailsExporter exporter = new GameDetailsExporter(gamesGenerated, outputFolder, resourcePath);
        exporter.ExportClues();
        exporter.ExportGameLegend();
        Console.WriteLine("Generated game");
       

    }
}