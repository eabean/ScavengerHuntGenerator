namespace ScavengerHuntGenerator
{
    public class Game
    {
        public string gameId;
        public List<Clue> clueList = new List<Clue>();
        public List<Location> selectedLocations;
        public List<Question> selectedQuestions;

        private readonly GameDetailsRepository _detailsRepository;
        private readonly GameSettings _settings;

        public Game(string gameId, GameDetailsRepository detailsRepository, GameSettings settings)
        {
            this.gameId = gameId;
            _detailsRepository = detailsRepository;
            _settings = settings;
        }

        public void GenerateGame()
        {
            var allLocations = _detailsRepository.ParseLocations();
            var allQuestions = _detailsRepository.ParseQuestions();
            var allFakeLocations = _detailsRepository.ParseFakeLocations();

            int fakeLocationsRequired = _settings.NumOfClues * (_settings.NumOfAnswers - 1);
            if (allFakeLocations.Count < fakeLocationsRequired)
                throw new Exception($"Not enough distinct fake locations to support {_settings.NumOfClues} clues. You need (number of clues) * (number of answers per clue-1) fake locations.");

            selectedLocations = RandomizeList(allLocations, _settings.NumOfClues);
            selectedQuestions = RandomizeList(allQuestions, _settings.NumOfClues);

            MapAnswersToLocations(selectedLocations, selectedQuestions, allFakeLocations);

            for (int i = 0; i < _settings.NumOfClues - 1; i++)
            {
                var clue = new Clue { location = selectedLocations[i], question = selectedQuestions[i] };
                clueList.Add(clue);
                Console.WriteLine($"Clue {i}: Location: {clue.location.locId}, {clue.location.decodedDescription}" +
                    $" Question: {clue.question.qId}, {clue.question.qText}");
            }
        }

        private void MapAnswersToLocations(List<Location> locations, List<Question> questions, List<Location> fakeLocations)
        {
            var remainingFakeLocations = fakeLocations;
            for (int i = 0; i < questions.Count; i++)
            {
                var fakeLocationSet = GetFakeLocationSet(remainingFakeLocations, _settings.NumOfAnswers - 1);
                remainingFakeLocations = remainingFakeLocations.Where(l => !fakeLocationSet.Any(f => f.locId == l.locId)).ToList();

                var mapping = questions[i].answersLocationMapping;
                var correctClueDescription = locations[i].clueDescription;
                var fakeIndex = 0;

                foreach (var answer in questions[i].qAnswers)
                {
                    mapping.Add(answer.anText, answer.isCorrect
                        ? correctClueDescription
                        : fakeLocationSet[fakeIndex++].clueDescription);
                }
            }
        }

        private List<Location> GetFakeLocationSet(List<Location> locations, int count)
        {
            List<Location> result;
            do { result = RandomizeList(locations, count); }
            while (result.Select(l => l.locId).Distinct().Count() != count);
            return result;
        }

        private List<T> RandomizeList<T>(List<T> list, int count)
        {
            return list.OrderBy(_ => Random.Shared.Next()).Take(count).ToList();
        }
    }
}
