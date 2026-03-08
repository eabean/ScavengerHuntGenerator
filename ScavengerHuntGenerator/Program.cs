using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using ScavengerHuntGenerator;


class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.License.SetNonCommercialPersonal("ScavengerHuntGenerator");
        var configuration = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        var settings = configuration.GetSection("GameSettings").Get<GameSettings>()
            ?? throw new Exception("GameSettings section missing from appsettings.json");

        var executableDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var projectDirectory = settings.ProjectDirectory
            ?? Directory.GetParent(executableDirectory)?.FullName
            ?? throw new Exception("Could not determine project directory.");

        var gameIds = Enumerable.Range('A', 26).Select(x => ((char)x).ToString()).ToList();

        var resourcesFolder = Path.Combine(projectDirectory, settings.ResourcesFolderName);
        var outputFolder = Path.Combine(projectDirectory, settings.OutputFolderName);
        var questionsDbPath = Path.Combine(resourcesFolder, settings.QuestionsDatabaseFileName);
        var clueTemplatePath = Path.Combine(resourcesFolder, settings.ClueTemplateFileName);

        GameDetailsRepository gameDetailsRepository = new GameDetailsRepository(questionsDbPath);

        List<Game> gamesGenerated = new List<Game>();
        for (int i = 0; i < settings.NumOfGames; i++)
        {
            var game = new Game(gameIds[i], gameDetailsRepository, settings);
            game.GenerateGame();
            gamesGenerated.Add(game);
        }

        GameDetailsExporter exporter = new GameDetailsExporter(gamesGenerated, outputFolder, clueTemplatePath, gameDetailsRepository, settings);
        exporter.ExportClues();
        exporter.ExportGameLegend();
        Console.WriteLine($"Generated game in {projectDirectory}");
    }
}
