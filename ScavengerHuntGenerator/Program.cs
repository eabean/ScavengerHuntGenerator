using ScavengerHuntGenerator;
using System;


class Program
{

    public List<string> gameIds = Enumerable.Range('A', 26).Select(x => ((char)x).ToString()).ToList();
    static void Main(string[] args)
    {
        var executableDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var projectDirectory = Directory.GetParent(executableDirectory).FullName;

        var resourcesFolderName = "Resources";
        var resourcesFolderPath = Path.Combine(projectDirectory, resourcesFolderName);
        var dbPath = resourcesFolderPath + @"\ScavengerHuntDb.xlsx";

        var outputFolderName = "Output";
        string outputFolder = Path.Combine(projectDirectory, outputFolderName);
        string resourcePath = resourcesFolderPath + @"\Clues2x2.docx";

        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(dbPath);
        var game = new Game("A", gameDetailsRepository);
        game.GenerateGame();
        List<Game> gamesGenerated = new List<Game> { game };
        GameDetailsExporter exporter = new GameDetailsExporter(gamesGenerated, outputFolder, resourcePath);
        exporter.ExportClues(game);
        exporter.ExportGameLegend();
        Console.WriteLine("Generated game");
       

    }
}