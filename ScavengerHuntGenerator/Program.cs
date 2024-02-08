using ScavengerHuntGenerator;
using System;


class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Users\Enid\Documents\Dev\ScavengerHuntDb.xlsx";
        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(filePath);
        var game = new Game(gameDetailsRepository);
        game.GenerateGame();

        Console.WriteLine("Generated game");
       

    }
}