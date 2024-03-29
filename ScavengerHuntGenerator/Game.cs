﻿using System;
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
        public const int NUM_OF_CLUES= 10;
        public const int NUM_OF_GAMES = 8;
        public const int NUM_OF_ANS = 4;
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
            var allFakeLocations = _detailsRepository.ParseFakeLocations();
            selectedLocations = RandomizeList(allLocations, NUM_OF_CLUES);
            selectedQuestions = RandomizeList(allQuestions, NUM_OF_CLUES);

            if(allFakeLocations.Count < NUM_OF_CLUES*(NUM_OF_ANS-1)) { throw new Exception($"Not enough distinct fake locations to support {NUM_OF_CLUES} clues."); }

            MapAnswersToLocations(selectedLocations, selectedQuestions, allFakeLocations);
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

        public async void MapAnswersToLocations(List<Location> locations, List<Question> questions, List<Location> fakeLocations)
        {
            var mutatingFakeLocations = fakeLocations;
            for(int i = 0;i < questions.Count;i++)
            {
                var qAnswers = questions[i].qAnswers;
                var fakeLocationAnswers = GetFakeLocationSet(mutatingFakeLocations, NUM_OF_ANS - 1);
                mutatingFakeLocations = mutatingFakeLocations.Where(l1 => !fakeLocationAnswers.Any(l2 => l2.locId == l1.locId)).ToList();

                var fakeLocationIndex = 0;
                foreach (var answer in qAnswers)
                {
                    var questionMapping = questions[i].answersLocationMapping;
                    var correctLocation = locations[i].clueDescription;
                    if (answer.isCorrect) 
                    {
                       questionMapping.Add(answer.anText, correctLocation);

                    } else
                    {
                        questionMapping.Add(answer.anText, fakeLocationAnswers[fakeLocationIndex].clueDescription);
                        fakeLocationIndex++;
                    }
                }

            }
        }

        public List<Location> GetFakeLocationSet(List<Location> originalList, int length)
        {
            var randomFakeLocations = RandomizeList(originalList, length);
            while (randomFakeLocations.Select(l => l.locId).Distinct().Count() != randomFakeLocations.Count)
            {
                randomFakeLocations = RandomizeList(originalList, length);
            }
            return randomFakeLocations;
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
