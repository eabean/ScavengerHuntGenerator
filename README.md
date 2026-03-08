# Scavenger Hunt Generator

Generates printable clue card documents and an answer key spreadsheet for a multi-team scavenger hunt. Each team receives a uniquely randomised set of clues and locations.

---

## How to Run

```bash
dotnet run --project ScavengerHuntGenerator/ScavengerHuntGenerator.csproj
```

Output files are written to the folder configured by `OutputFolderName` (default: the `Output` folder inside your `ProjectDirectory`).

---

## Configuration (`appsettings.json`)

All settings live in `ScavengerHuntGenerator/appsettings.json` under the `GameSettings` key.

### Paths

| Setting | Description |
|---|---|
| `ProjectDirectory` | Absolute path to the folder that contains the `Resources` and `Output` subfolders. If omitted, defaults to one level above the build output directory. |
| `ResourcesFolderName` | Name of the subfolder containing input files. Default: `Resources` |
| `OutputFolderName` | Name of the subfolder where generated files are written. Default: `Output` |
| `QuestionsDatabaseFileName` | Name of the Excel database file inside `Resources`. Default: `ScavengerHuntDbShared.xlsx` |
| `ClueTemplateFileName` | Name of the Word `.docx` template used for clue cards. Default: `Clues2x2.docx` |

### Game Rules

| Setting | Description |
|---|---|
| `NumOfGames` | Number of simultaneous games (teams) to generate. Max 26 (labelled A–Z). Default: `8` |
| `NumOfClues` | Number of clue cards per game. Default: `10` |
| `NumOfAnswers` | Number of answer choices per question (1 correct + rest fake locations). Default: `4` |
| `MaxGames` | Upper limit on supported games. Default: `26` |
| `MaxClues` | Upper limit on supported clues per game. Default: `12` |

### Content

| Setting | Description |
|---|---|
| `FinalInstruction` | Text printed on the last clue card, shown after all questions are answered. |

---

## Resources Folder

Place the following files in your `Resources` folder before running:

### `ScavengerHuntDbShared.xlsx` (or your configured database filename)

An Excel workbook with **three worksheets** in this order:

**Sheet 1 — Questions**

| Column | Content |
|---|---|
| A | Question ID (integer) |
| B | Question text |
| C | Answer choices, one per line. Mark the correct answer by appending `*` to the end of it (e.g. `Tower Bridge*`). |

**Sheet 2 — Locations**

| Column | Content |
|---|---|
| A | Location ID (short unique identifier, e.g. `L1`) |
| B | Decoded description (human-readable name, e.g. `Reception Desk`) |
| C | Clue description (cryptic hint printed on the card, e.g. `Where you first say hello`) |

**Sheet 3 — Fake Locations**

| Column | Content |
|---|---|
| A | Location ID |
| B | Clue description (used as wrong-answer destinations) |

The number of fake locations must be at least `NumOfClues × (NumOfAnswers − 1)` to ensure each the game has enough unique wrong answers.

### `Clues2x2.docx` (or your configured template filename)

A Word document containing a single table. The table must have at least `NumOfClues + 1` cells — one per question plus one for the final instruction card. The generator overwrites each cell with the clue content.

---

## Reading the Game Legend (`GameLegend.xlsx`)

The legend is the answer key for all games. It is written to your `Output` folder.

### Clue Grid (columns B onward)

Each game occupies two consecutive rows. For a game labelled `A`:

| Row | Content |
|---|---|
| `Game A Location` | The Location ID where each clue card is physically placed |
| `Game A Answer` | The correct answer letter for the question on each clue card |

The columns are headed **Clue 1, Clue 2, … Clue N+1** (one more than `NumOfClues`):

- **Clue 1** — no Location (the first card is handed to players at the start); has an Answer
- **Clue 2 to Clue N** — has both a Location and an Answer
- **Clue N+1** — has a Location (where the final card is placed); no Answer (final instruction only)

Location cells are **colour-coded**: each unique Location ID has a consistent colour across all games and matches the colour in the location reference table on the right.

### Location Reference Table (far right)

Lists all possible correct locations with three columns: `LocationId`, `decodedDescription`, and `clueDescription`. The `LocationId` cells are colour-coded to match the clue grid.
