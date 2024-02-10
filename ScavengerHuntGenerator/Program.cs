using ScavengerHuntGenerator;
using System;


class Program
{
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
        var game = new Game(gameDetailsRepository);
        game.GenerateGame();
        GameDetailsExporter exporter = new GameDetailsExporter(outputFolder, resourcePath);
        exporter.ExportClues();
        Console.WriteLine("Generated game");
       

    }
}