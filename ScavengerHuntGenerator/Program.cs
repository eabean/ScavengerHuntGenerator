using ScavengerHuntGenerator;
using System;


class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Users\Enid\Documents\Dev\ScavengerHuntDb.xlsx";
        string exportPath= @"C:\Users\Enid\Documents\Dev\ScavengerOutput\game.doc";
        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(filePath);
        var game = new Game(gameDetailsRepository);
        game.GenerateGame();
        GameDetailsExporter exporter = new GameDetailsExporter(exportPath);
        exporter.CreateWordDocument();
        Console.WriteLine("Generated game");
       

    }
}