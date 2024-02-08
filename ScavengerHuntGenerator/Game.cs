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
    }

    public class Answer
    {
        //public int anId;
        public string anText;
        public bool isCorrect;
    }

    public class Question
    {
        public int qId;
        public string qText;
        public List<Answer> qAnswers;
        public Dictionary<Answer, Location> answersLocationMapping;

        public Question()
        {
                
        }

        public Question(int id, string text, List<Answer> answers)
        {
            this.qId = id;   
            this.qText = text;
            this.qAnswers = answers;
        }

        public void MapAnswersToLocations(Question question, List<Location> locations, Location correctLocation)
        {
            return;

        }

    }

    public class Clue
    {
        public Location location;
        public Question question;
    }


    public class Game
    {
        public List<Clue> clueList;
        public const int MAX_CLUES = 10;

        public Game() { }


        public void GenerateGame()
        {
            clueList = new List<Clue>();
        }
    }
}
