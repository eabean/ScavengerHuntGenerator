using ScavengerHuntGenerator;
using System;


class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Users\Enid\Documents\Dev\ScavengerHuntDb.xlsx";
        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(filePath);
        var questions = gameDetailsRepository.ParseQuestions();
        var locations = gameDetailsRepository.ParseLocations();

       

    }
}