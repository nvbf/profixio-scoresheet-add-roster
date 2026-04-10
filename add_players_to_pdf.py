"""
Add player lists to tournament schedule PDF.

This script reads player data from players.xlsx and adds player rosters
to the kampskjema PDF, creating a new PDF with players listed
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
from PyPDF2 import PdfReader, PdfWriter
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

    if 'Surname' in df.columns:
        required_cols = ['Team', 'Class', 'Number', 'Name', 'Surname', 'Rolle']
    else:
        required_cols = ['Sp lag', 'Klasse', 'Draktnr', 'Fornavn', 'Etternavn', 'Rolle']

    # Verify expected columns exist
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns in Excel file: {missing_cols}")

    team_col = required_cols[0]
    class_col = required_cols[1]
    number_col = required_cols[2]
    name_col = required_cols[3]
    surname_col = required_cols[4]
    role_col = required_cols[5]

    # Group players by (Team, Class)
    player_dict = defaultdict(list)

    for _, row in df.iterrows():
        if str(row[role_col]).strip().lower() != 'spiller':
            continue  # Skip non-player roles

        team = str(row[team_col]).strip()
        class_name = str(row[class_col]).strip()

        # Convert number to int, handling various formats
        try:
            number = int(float(row[number_col]))  # Handle both int and float strings
        except (ValueError, TypeError):
            number = 0

        name = str(row[name_col]).strip()
        surname = str(row[surname_col]).strip()

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
                    if top_words[i]['x0'] - top_words[i - 1]['x1'] > 50:
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
                team_class = None
                for word in words:
                    if 'GU15' in word['text']:
                        team_class = 'GU15'
                        break
                    elif 'JU15' in word['text']:
                        team_class = 'JU15'
                        break

                # Keep the class as-is (GU15 or JU15 to match Excel)
                if team_class is None:
                    print(f"  Page {page_num + 1}: Could not identify class")
                    continue

                # Calculate positions for player lists
                # Use the leftmost word's x for each team
                team1_x = teams[0][0]['x0']
                team2_x = teams[1][0]['x0']

                # Players should start below "Nr" header
                team1_coords = {
                    'x': team1_x - 43,     # Adjust X to align with pre-printed "Nr" column
                    'y': page.height - 62  # Adjust Y to align with first line of players
                }
                team2_coords = {
                    'x': team2_x - 45,     # Adjust X to align with pre-printed "Nr" column
                    'y': page.height - 62  # Adjust Y to align with first line of players
                }

                teams_info.append((
                    page_num,
                    team1_name, team_class, team1_coords,
                    team2_name, team_class, team2_coords
                ))

            except Exception as e:
                print(f"  Page {page_num + 1}: Error extracting info - {e}")

    return teams_info


def create_player_overlay(page_width, page_height, players, x, y, font_size=8, team_side='left'):
    """
    Create a PDF overlay with player list at specified coordinates.

    Args:
        page_width: Width of the page
        page_height: Height of the page
        players: List of (number, name, surname) tuples
        x: X coordinate for text start
        y: Y coordinate for text start (from bottom)
        font_size: Font size for text
        team_side: 'left' for team 1 or 'right' for team 2 (determines white box position)

    Returns:
        io.BytesIO: PDF overlay as bytes
    """
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    can.setFont("Helvetica", font_size)

    fix_libero_info(can, x, y, font_size, team_side)

    # Starting Y position (PDF coordinates are from bottom)
    current_y = y
    line_height = font_size + 2

    # Add each player
    current_line = 1
    for number, name, surname in players[:MAX_PLAYERS_PER_TEAM]:
        if number > 0:
            can.drawString(x, current_y, f"{number:>2}")
        player_text = f"{name} {surname}"
        can.drawString(x + 40, current_y, player_text)
        current_y -= line_height
        # Add som extra spacing after to match pre-printed lines
        if current_line in [8, 9, 10, 11]:
            current_y -= 1
        if current_line in [12]:
            current_y -= 2
        current_line += 1

    can.save()
    packet.seek(0)
    return packet

def fix_libero_info(can, x, y, font_size, team_side):
    """
    Adds a white box to cover "Libero" labels next to lines 13-14 and
    replace "Lisens" column header with "Libero".

    Args:
        can: Canvas object to draw on
        x: Current X coordinate on canvas
        y: Current Y coordinate on canvas (starting point for player list)
        font_size: Font size for text
        team_side: 'left' for team 1 or 'right' for team 2 (determines white box position)

    """
    can.setFillColorRGB(1, 1, 1)  # White background
    can.setStrokeColorRGB(1, 1, 1)  # White border

    # Draw white box to cover "Libero" labels next to lines 13-14.
    # =====
    line_height = font_size + 2
    # Calculate Y position for line 13 (accounting for spacing after lines 8-12)
    line_13_y = y - (12 * line_height) - 1 - 1 - 1 - 1 - 2  # 12 lines + extra spacing
    box_height = 25  # Height to cover lines 13-14
    box_width = 27   # Width to cover the area next to line 13-14

    # Position box based on team side
    if team_side == 'left':
        # Team 1: box on the left side of roster
        box_x = x - 36
    else:
        # Team 2: box on the right side of roster
        box_x = x + 158

    can.rect(box_x, line_13_y - box_height + line_height, box_width, box_height, fill=1, stroke=1)

    # Draw "Libero" box above roster hiding pre-printed "Lisens" column header.
    # =====
    box_height = 8
    box_width = 22
    box_x = x + 12
    box_y = y + 9 # Position above the roster

    can.rect(box_x, box_y, box_width, box_height, fill=1, stroke=1)

    can.setFillColorRGB(0, 0, 0)
    can.drawString(box_x, box_y, "Libero")

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

    skipped_rosters = []
    processed_rosters = 0
    pages_without_matches = []

    # Process each page
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]

        if page_num in page_teams:
            team_info = page_teams[page_num]
            team1_name, team1_class, team1_coords, team2_name, team2_class, team2_coords = team_info

            # Get page dimensions
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)

            # Track if teams were found in Excel data
            team1_found = False
            team2_found = False

            # Process team 1 (case-insensitive lookup)
            team1_lower = (team1_name.lower(), team1_class.lower())
            players_team1 = []
            if team1_lower in case_map:
                team1_found = True
                team1_key = case_map[team1_lower]
                players = player_data[team1_key]
                if len(players) > MAX_PLAYERS_PER_TEAM:
                    print(
                        f"  Page {page_num + 1}: WARNING: {team1_name} ({team1_class}) has {len(players)} players (max {MAX_PLAYERS_PER_TEAM}). Skipping player names for this team.")
                    skipped_rosters.append(f"{team1_name} ({team1_class})")
                    players_team1 = []  # Empty list, but still draw white boxes
                else:
                    players_team1 = players
                    processed_rosters += 1

            # Always create overlay for team 1 (includes white boxes even if no players)
            overlay_packet = create_player_overlay(
                page_width, page_height, players_team1,
                team1_coords['x'], team1_coords['y'],
                team_side='left'
            )
            overlay_page = PdfReader(overlay_packet).pages[0]
            page.merge_page(overlay_page)

            # Process team 2 (case-insensitive lookup)
            team2_lower = (team2_name.lower(), team2_class.lower())
            players_team2 = []
            if team2_lower in case_map:
                team2_found = True
                team2_key = case_map[team2_lower]
                players = player_data[team2_key]
                if len(players) > MAX_PLAYERS_PER_TEAM:
                    print(
                        f"  Page {page_num + 1}: WARNING: {team2_name} ({team2_class}) has {len(players)} players (max {MAX_PLAYERS_PER_TEAM}). Skipping player names for this team.")
                    skipped_rosters.append(f"{team2_name} ({team2_class})")
                    players_team2 = []  # Empty list, but still draw white boxes
                else:
                    players_team2 = players
                    processed_rosters += 1

            # Always create overlay for team 2 (includes white boxes even if no players)
            overlay_packet = create_player_overlay(
                page_width, page_height, players_team2,
                team2_coords['x'], team2_coords['y'],
                team_side='right'
            )
            overlay_page = PdfReader(overlay_packet).pages[0]
            page.merge_page(overlay_page)

            # Report if not both teams were found
            if not (team1_found and team2_found):
                missing = []
                if not team1_found:
                    missing.append(f"{team1_name} ({team1_class})")
                if not team2_found:
                    missing.append(f"{team2_name} ({team2_class})")
                pages_without_matches.append((page_num + 1, missing))

        writer.add_page(page)

    # Write the output PDF
    with open(output_pdf, 'wb') as output_file:
        writer.write(output_file)

    print(f"\nProcessing complete!")
    print(f"  Rosters processed: {processed_rosters}")
    print(f"  Rosters skipped (>{MAX_PLAYERS_PER_TEAM} players): {len(skipped_rosters)}")

    if pages_without_matches:
        print(f"\nPages without matching teams in Excel:")
        for page_num, missing_teams in pages_without_matches:
            print(f"  Page {page_num}: {', '.join(missing_teams)}")

    print(f"\nOutput saved to: {output_pdf}")


def main():
    global MAX_PLAYERS_PER_TEAM
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
    parser.add_argument('--max-players', type=int, default=MAX_PLAYERS_PER_TEAM,
                        help=f'Maximum players per team (default: {MAX_PLAYERS_PER_TEAM})')

    args = parser.parse_args()

    # Update max players from arguments
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
