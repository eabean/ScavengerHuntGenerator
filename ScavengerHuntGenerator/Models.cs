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
        public string anText;
        public bool isCorrect;
    }

    public class Question
    {
        public int qId;
        public string qText;
        public List<Answer> qAnswers;
        public Dictionary<string, string> answersLocationMapping = new Dictionary<string, string>();

        public Question() { }

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
}
