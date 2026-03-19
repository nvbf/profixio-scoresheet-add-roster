"""
Add player lists to tournament schedule PDF.

This script reads player data from players.xlsx and adds player rosters
to kampskjema-fredag-puljespill.pdf, creating a new PDF with players listed
below each team name.

Requirements:
- pandas
- openpyxl
- pdfplumber
- reportlab
- PyPDF2
"""

import pandas as pd
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter, Transformation
import io
import sys
import argparse
from pathlib import Path
from collections import defaultdict

# Configuration
MAX_PLAYERS_PER_TEAM = 14


def load_player_data(excel_path):
    """
    Load player data from Excel file and organize by team and class.

    Returns:
        dict: {(team_name, class): [(number, name, surname), ...]}
    """
    print(f"Loading player data from {excel_path}...")
    df = pd.read_excel(excel_path)

    # Verify expected columns exist
    required_cols = ['Team', 'Class', 'Number', 'Name', 'Surname']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns in Excel file: {missing_cols}")

    # Group players by (Team, Class)
    player_dict = defaultdict(list)

    for _, row in df.iterrows():
        team = str(row['Team']).strip()
        class_name = str(row['Class']).strip()

        # Convert number to int, handling various formats
        try:
            number = int(float(row['Number']))  # Handle both int and float strings
        except (ValueError, TypeError):
            number = 0

        name = str(row['Name']).strip()
        surname = str(row['Surname']).strip()

        key = (team, class_name)
        player_dict[key].append((number, name, surname))

    # Sort players by number within each team
    for key in player_dict:
        player_dict[key].sort(key=lambda x: x[0])

    # Create a case-insensitive lookup dictionary
    # Map lowercase (team, class) -> original (team, class) key
    case_insensitive_map = {}
    for (team, class_name) in player_dict.keys():
        lower_key = (team.lower(), class_name.lower())
        case_insensitive_map[lower_key] = (team, class_name)

    print(f"Loaded {len(player_dict)} teams with player data")
    return dict(player_dict), case_insensitive_map


def normalize_text(text):
    """Fix common character encoding issues in PDF extraction."""
    # Common encoding issues when extracting from PDF
    replacements = {
        '°': 'ø',
        'σ': 'å',
        '╪': 'Ø',
        'Σ': 'Å',
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text


def extract_team_info_from_pdf(pdf_path):
    """
    Extract team names and their positions from the PDF.

    Returns:
        list: [(page_num, team1_name, team1_class, team1_coords, team2_name, team2_class, team2_coords), ...]
    """
    print(f"Analyzing PDF structure: {pdf_path}...")
    teams_info = []

    with pdfplumber.open(pdf_path) as pdf:
        print(f"Total pages: {len(pdf.pages)}")

        for page_num, page in enumerate(pdf.pages):
            try:
                # Extract words with coordinates
                words = page.extract_words()

                if not words:
                    print(f"  Page {page_num + 1}: No text found, skipping")
                    continue

                # Find team names at the top of the page (y between 15-35 from bottom)
                top_words = [w for w in words if 15 <= w['top'] <= 35]

                if len(top_words) < 2:
                    print(f"  Page {page_num + 1}: Not enough words at top")
                    continue

                # Sort by x position
                top_words.sort(key=lambda w: w['x0'])

                # Group words by proximity in X direction (gap > 100 points means new team)
                teams = []
                current_team = [top_words[0]]

                for i in range(1, len(top_words)):
                    # If the gap between this word and the previous one is large, it's a new team
                    if top_words[i]['x0'] - top_words[i - 1]['x0'] > 100:
                        teams.append(current_team)
                        current_team = [top_words[i]]
                    else:
                        current_team.append(top_words[i])

                # Add the last team
                teams.append(current_team)

                # We should have exactly 2 teams
                if len(teams) != 2:
                    print(f"  Page {page_num + 1}: Found {len(teams)} teams instead of 2")
                    continue

                # Combine words in each team to form team names
                team1_name = ' '.join([w['text'].strip() for w in teams[0] if len(w['text'].strip()) > 0])
                team2_name = ' '.join([w['text'].strip() for w in teams[1] if len(w['text'].strip()) > 0])

                # Normalize encoding issues
                team1_name = normalize_text(team1_name)
                team2_name = normalize_text(team2_name)

                if not team1_name or not team2_name:
                    print(f"  Page {page_num + 1}: Could not extract both team names")
                    continue

                # Find class designation (GU15 or JU15)
                class_word = None
                for word in words:
                    if 'GU15' in word['text']:
                        class_word = 'GU15'
                        break
                    elif 'JU15' in word['text']:
                        class_word = 'JU15'
                        break

                # Keep the class as-is (GU15 or JU15 to match Excel)
                if class_word == 'GU15':
                    team_class = 'GU15'
                elif class_word == 'JU15':
                    team_class = 'JU15'
                else:
                    print(f"  Page {page_num + 1}: Could not identify class")
                    continue

                # Calculate positions for player lists
                # Use the leftmost word's x for each team
                team1_x = teams[0][0]['x0']
                team2_x = teams[1][0]['x0']

                # Players should start below "Nr" header
                # Nr is at 2cm from top (≈57 points), first player row is 14 points below
                team1_coords = {
                    'x': team1_x,
                    'y': page.height - 71  # 71 points from top edge (2cm + 5mm)
                }
                team2_coords = {
                    'x': team2_x,
                    'y': page.height - 71  # 71 points from top edge (2cm + 5mm)
                }

                teams_info.append((
                    page_num,
                    team1_name, team_class, team1_coords,
                    team2_name, team_class, team2_coords
                ))
                print(f"  Page {page_num + 1}: {team1_name} ({team_class}) vs {team2_name} ({team_class})")

            except Exception as e:
                print(f"  Page {page_num + 1}: Error extracting info - {e}")

    return teams_info


def create_player_overlay(page_width, page_height, team_name, team_class, players, x, y, font_size=8):
    """
    Create a PDF overlay with player list at specified coordinates.

    Args:
        page_width: Width of the page
        page_height: Height of the page
        team_name: Name of the team
        team_class: Class (J15 or G15)
        players: List of (number, name, surname) tuples
        x: X coordinate for text start
        y: Y coordinate for text start (from bottom)
        font_size: Font size for text

    Returns:
        io.BytesIO: PDF overlay as bytes
    """
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    can.setFont("Helvetica", font_size)

    # Starting Y position (PDF coordinates are from bottom)
    current_y = y + 9
    line_height = 10  # 5mm spacing to match pre-printed roster lines

    # Add each player
    current_line = 1
    for number, name, surname in players[:MAX_PLAYERS_PER_TEAM]:
        if number > 0:
            can.drawString(x - 45, current_y, f"{number:>2}")
        player_text = f"{name} {surname}"
        can.drawString(x - 5, current_y, player_text)
        current_y -= line_height
        if current_line in [9, 11, 12]:
            current_y -= 2  # Extra spacing after 9 and 12 players to match pre-printed lines
        current_line += 1

    can.save()
    packet.seek(0)
    return packet


def add_players_to_pdf(input_pdf, output_pdf, player_data, case_map, teams_info):
    """
    Create a new PDF with player lists added to each page.

    Args:
        input_pdf: Path to input PDF
        output_pdf: Path to output PDF
        player_data: Dictionary of player data by (team, class)
        case_map: Case-insensitive mapping of team names
        teams_info: List of team information extracted from PDF
    """
    print("\nProcessing PDF and adding player lists...")

    # Read the input PDF
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # Create a mapping of page_num to team info for quick lookup
    page_teams = {}
    for info in teams_info:
        page_num = info[0]
        page_teams[page_num] = info[1:]  # Everything except page_num

    skipped_teams = []
    processed_teams = 0

    # Process each page
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]

        if page_num in page_teams:
            team_info = page_teams[page_num]
            team1_name, team1_class, team1_coords, team2_name, team2_class, team2_coords = team_info

            # Get page dimensions
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)

            # Process team 1 (case-insensitive lookup)
            team1_lower = (team1_name.lower(), team1_class.lower())
            if team1_lower in case_map:
                team1_key = case_map[team1_lower]
                players = player_data[team1_key]
                if len(players) > MAX_PLAYERS_PER_TEAM:
                    print(
                        f"  Page {page_num + 1}: WARNING: {team1_name} ({team1_class}) has {len(players)} players (max {MAX_PLAYERS_PER_TEAM}). Skipping this team.")
                    skipped_teams.append(f"{team1_name} ({team1_class})")
                else:
                    # Create overlay for team 1
                    overlay_packet = create_player_overlay(
                        page_width, page_height,
                        team1_name, team1_class, players,
                        team1_coords['x'] + 2, team1_coords['y']
                    )
                    overlay_page = PdfReader(overlay_packet).pages[0]
                    page.merge_page(overlay_page)
                    processed_teams += 1

            # Process team 2 (case-insensitive lookup)
            team2_lower = (team2_name.lower(), team2_class.lower())
            if team2_lower in case_map:
                team2_key = case_map[team2_lower]
                players = player_data[team2_key]
                if len(players) > MAX_PLAYERS_PER_TEAM:
                    print(
                        f"  Page {page_num + 1}: WARNING: {team2_name} ({team2_class}) has {len(players)} players (max {MAX_PLAYERS_PER_TEAM}). Skipping this team.")
                    skipped_teams.append(f"{team2_name} ({team2_class})")
                else:
                    # Create overlay for team 2
                    overlay_packet = create_player_overlay(
                        page_width, page_height,
                        team2_name, team2_class, players,
                        team2_coords['x'], team2_coords['y']
                    )
                    overlay_page = PdfReader(overlay_packet).pages[0]
                    page.merge_page(overlay_page)
                    processed_teams += 1

        writer.add_page(page)

    # Write the output PDF
    with open(output_pdf, 'wb') as output_file:
        writer.write(output_file)

    print(f"\nProcessing complete!")
    print(f"  Teams processed: {processed_teams}")
    print(f"  Teams skipped (>{MAX_PLAYERS_PER_TEAM} players): {len(skipped_teams)}")
    print(f"\nOutput saved to: {output_pdf}")


def main():
    """Main execution function."""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description='Add player lists to tournament schedule PDF',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''Examples:
  python add_players_to_pdf.py input.pdf players.xlsx output.pdf
  python add_players_to_pdf.py schedule.pdf roster.xlsx result.pdf
        '''
    )
    parser.add_argument('input_pdf', help='Input PDF file with tournament schedule')
    parser.add_argument('input_excel', help='Excel file with player data')
    parser.add_argument('output_pdf', help='Output PDF file to create')
    parser.add_argument('--max-players', type=int, default=14,
                        help='Maximum players per team (default: 14)')

    args = parser.parse_args()

    # Update max players from arguments
    global MAX_PLAYERS_PER_TEAM
    MAX_PLAYERS_PER_TEAM = args.max_players

    print("=" * 80)
    print("PDF Player List Generator")
    print("=" * 80)

    # Convert to Path objects
    input_pdf = Path(args.input_pdf)
    input_excel = Path(args.input_excel)
    output_pdf = Path(args.output_pdf)

    # Check input files exist
    if not input_pdf.exists():
        print(f"ERROR: Input PDF not found: {input_pdf}")
        sys.exit(1)

    if not input_excel.exists():
        print(f"ERROR: Input Excel file not found: {input_excel}")
        sys.exit(1)

    # If output file exists, rename it to .old
    if output_pdf.exists():
        old_file = output_pdf.with_suffix(output_pdf.suffix + '.old')
        print(f"Output file exists. Renaming {output_pdf.name} to {old_file.name}")
        # Remove .old file if it exists
        if old_file.exists():
            old_file.unlink()
        output_pdf.rename(old_file)

    try:
        # Step 1: Load player data
        player_data, case_map = load_player_data(str(input_excel))

        # Step 2: Extract team information from PDF
        teams_info = extract_team_info_from_pdf(str(input_pdf))

        if not teams_info:
            print("\nERROR: No team information could be extracted from PDF.")
            print("Please verify the PDF structure and update the extraction logic.")
            sys.exit(1)

        # Step 3: Add players to PDF
        add_players_to_pdf(str(input_pdf), str(output_pdf), player_data, case_map, teams_info)

        print("\n" + "=" * 80)
        print("SUCCESS! Check the output file and verify the results.")
        print("=" * 80)

    except Exception as e:
        print(f"\nERROR: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
