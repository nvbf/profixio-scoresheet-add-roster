# PDF Player List Generator for Profixio scoresheet

This script adds player rosters to tournament scoresheet PDFs (exported from Profixio).
It reads player data from an Excel file with players (export from Profixio) and overlays the player lists onto the PDF at specified positions.

Since the players are sorted after number, the Lisens column is renamed to Libero and user for libero assignment.
This is particularly useful for U15 NM where more than two players can be libero.

## Requirements

- Python 3.10 or higher
- Required Python packages (see `requirements.txt`)
- Able to use the command line

Everything is tested on Windows using PowerShell to ensure that Linux / Mac isn't required.
Commands below can be pasted into PowerShell.

## Installation

1. Install Python. On Linux use your package manager, on Windows use [Python Install Manager](https://apps.microsoft.com/detail/9nq7512cxl7t).
2. Downloed the latest release of the script from the [release page](https://github.com/nvbf/profixio-scoresheet-add-roster/releases).
3. Unpack the release and enter the new directory from your command line application.
4. Install Python dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Download data from Profixio

- Scoresheet (PDF): Export from "Resultatregistrering", choose "Konkurranseskjema 3setts" (and save as "kampskjema.pdf" in stead of the default ""kampkort").
- Players spreadsheet (EXcel): Export from "Spillere"

For convenience place the files in the same folder as the Python 

You need to have access to the tournament (and have logged in) to export these files.

### Quick start

From the "release" direcory, simply run the script:

```bash
python add_players_to_pdf.py [-h] [--max-players MAX_PLAYERS] input.pdf input.xlsx output.pdf
```

The script will:
1. Read player data from input Excel file - typically `players.xlsx` - exported from Profixio
2. Extract team information from input PDF with one scoresheets per page
3. Add player lists to the PDF (default max 14 players per team)
4. Save new scoresheets with players as a new output PDF file

### Input files

**Input scoresheet PDF File**
- Should contain team names and class designations
- Each page should have space for player lists below team names

**Input player list Excel file**
- Required columns: `Team`, `Class`, `Number`, `Name`, `Surname`
- Teams are matched by exact name and class (JU15 or GU15)
- Players are automatically sorted by number

### Output

- **New scoresheet PDF file**
- **Console output**: Summary of processed teams and any warnings

## Player Formatting

Players are formatted as: `{Number} {Name} {Surname}`

Example:
```
1 John Doe
2 Jane Smith
3 Bob Johnson
```

Aligning with Nr and Name columns in the scoresheet

## Teams with more than max players

If a team has more than 14 (default) players:
- A **warning message** is printed to the console
- The team is **skipped entirely** (no players added)
- This prevents overflow and maintains PDF layout integrity

## Troubleshooting

### Issue: Team names not found

Missing teams or non-matched teams in the scoresheet are reported.

### Adjusting positioning of player list

The text position is controlled in the `extract_team_info_from_pdf()` function - 

```python
team1_coords = {
    'x': team1_x,
    'y': page.height - 71
}
team2_coords = {
    'x': team2_x,
    'y': page.height - 71
}
```

Coordinate system:
- `x`: Distance from left edge (in points, 72 points = 1 inch)
- `y`: Distance from bottom edge of page (not top!)

### Adjusting font size and spacing

In the `create_player_overlay()` function:

```python
font_size = 8  # Font size in points (argument)
line_height = font_size + 2  # Space between lines
```

## Technical notes

**PDF coordinate system**: Unlike most graphics systems, PDF coordinates start at the bottom-left, not top-left. The `y` coordinate increases upward.

**Page Size**: The script assumes standard page sizes. If your PDF uses custom dimensions, the coordinates may need adjustment.

**Text extraction**: The script uses `pdfplumber` for text extraction, which works well with text-based PDFs but not with scanned images.
