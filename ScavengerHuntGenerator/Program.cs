using ScavengerHuntGenerator;
using System;


class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Users\Enid\Documents\Dev\ScavengerHuntDb.xlsx";
        string exportFolder= @"C:\Users\Enid\Documents\Dev\ScavengerOutput\";
        string resourcePath= @"C:\Users\Enid\Documents\Dev\Resources\Clues2x2.docx";
        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(filePath);
        var game = new Game(gameDetailsRepository);
        game.GenerateGame();
        GameDetailsExporter exporter = new GameDetailsExporter(exportFolder, resourcePath);
        exporter.ExportClues();
        Console.WriteLine("Generated game");
       

    }
}