namespace ScavengerHuntGenerator
{
    public class GameSettings
    {
        public int NumOfClues { get; set; } = 10;
        public int NumOfGames { get; set; } = 8;
        public int NumOfAnswers { get; set; } = 4;
        public int MaxGames { get; set; } = 26;
        public int MaxClues { get; set; } = 12;
        public string? ProjectDirectory { get; set; }
        public string ResourcesFolderName { get; set; } = "Resources";
        public string OutputFolderName { get; set; } = "Output";
        public string QuestionsDatabaseFileName { get; set; } = "ScavengerHuntDbShared.xlsx";
        public string ClueTemplateFileName { get; set; } = "Clues2x2.docx";
        public string FinalInstruction { get; set; } = "You've finished! Please return to the 27th floor kitchen and submit your envelopes to secure your placing.";
    }
}
