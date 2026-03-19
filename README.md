# PDF Player List Generator for Profixio scoresheet

This script adds player rosters to tournament scoresheet PDFs (exported from Profixio).
It reads player data from an Excel file with players (export from Profixio) and overlays the player lists onto the PDF at specified positions.

## Requirements

- Python 3.8 or higher
- Required Python packages (see `requirements.txt`)

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Quick start

Simply run the script:
```bash
python add_players_to_pdf.py [-h] [--max-players MAX_PLAYERS] input_pdf input_excel output_pdf
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
- Teams are matched by exact name and class (J15 or G15)
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