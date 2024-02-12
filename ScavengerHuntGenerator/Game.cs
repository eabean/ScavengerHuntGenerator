using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScavengerHuntGenerator
{

    public class Location
    {
        public string locId;
        public string decodedDescription;
        public string clueDescription;

        public Location()
        {
            
        }

    }

    public class Answer
    {
        public string anText;
        public bool isCorrect;
    }

    public class Question
    {
        public int qId;
        public string qText;
        public List<Answer> qAnswers;
        public Dictionary<string, string> answersLocationMapping = new Dictionary<string, string>();

        public Question()
        {
                
        }

        public Question(int id, string text, List<Answer> answers)
        {
            this.qId = id;   
            this.qText = text;
            this.qAnswers = answers;
        }
    }

    public class Clue
    {
        public Location location;
        public Question question;
    }


    public class Game
    {
        public string gameId;
        public List<Clue> clueList = new List<Clue>();
        public List<Location> selectedLocations;
        public List<Question> selectedQuestions;

        private GameDetailsRepository _detailsRepository;
        public const int NUM_OF_CLUES= 5;
        public const int NUM_OF_GAMES = 3;
        public const int MAX_GAMES = 26;
        public const int MAX_CLUES = 12;


        public Game(string gameId, GameDetailsRepository detailsRepository)
        {
            this.gameId = gameId;
            _detailsRepository = detailsRepository;
        }

        public void GenerateGame()
        {
            var allLocations = _detailsRepository.ParseLocations();
            var allQuestions = _detailsRepository.ParseQuestions();
            selectedLocations = RandomizeList(allLocations, NUM_OF_CLUES);
            selectedQuestions = RandomizeList(allQuestions, NUM_OF_CLUES);
            MapAnswersToLocations(selectedLocations, selectedQuestions);
            for(int i = 0; i < NUM_OF_CLUES-1; i++)
            {
                var c = new Clue();
                c.location = selectedLocations[i];
                c.question = selectedQuestions[i];
                clueList.Add(c);
                Console.WriteLine($"Clue {i}: Location: {c.location.locId.ToString()}, {c.location.decodedDescription.ToString()}" +
                    $" Question:  {c.question.qId}, {c.question.qText.ToString()}, {c.question.answersLocationMapping.ToString()}");
            }

        }

        public async void MapAnswersToLocations(List<Location> locations, List<Question> questions)
        {
            for(int i = 0;i < questions.Count;i++)
            {
                var qAnswers = questions[i].qAnswers;
                foreach(var answer in qAnswers)
                {
                    var questionMapping = questions[i].answersLocationMapping;
                    var correctLocation = locations[i].clueDescription;
                    if (answer.isCorrect) 
                    {
                       questionMapping.Add(answer.anText, correctLocation);
                    } else
                    {
                        questionMapping.Add(answer.anText, "fake location");
                    }
                }

            }
        }

        public List<T> RandomizeList<T>(List<T> originalList, int length)
        {
            var rand = new Random();
            var shuffledList = originalList.OrderBy(x => rand.Next()).ToList();
            List<T> selectedItems = shuffledList.Take(length).ToList();

            return selectedItems;
        }
    }
}
