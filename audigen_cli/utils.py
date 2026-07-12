
import math
from datetime import datetime, timedelta
import os
from pathlib import Path
from audigen_cli import config as cfg

DATE_FORMAT = "%d-%m-%Y"

def _validate_date(value: str) -> bool | str:
    try:
        datetime.strptime(value, DATE_FORMAT)
        return True
    except ValueError:
        return f"Expected date in format DD-MM-YYYY e.g. 20-04-2025, got: {value}"

def _validate_date_range(start: str , end: str) -> bool:
    return datetime.strptime(start, DATE_FORMAT) <= datetime.strptime(end, DATE_FORMAT)

def parse_string_to_dateTime(value: str| datetime):
    if isinstance(value,datetime):
        return value
    return datetime.strptime(value, DATE_FORMAT)
 
def resolve_output_dir(ticket_id: str, output_arg: str | None) -> Path:
    """
    Priority: --output arg > config default > current working directory.
    Always creates a subfolder named after the ticket.
    """
    # Sanitize ticket_id — reject absolute paths, traversal, nested segments
    ticket_path = Path(ticket_id)
    if (
        ticket_path.is_absolute() 
        or ".." in ticket_path.parts 
        or len(ticket_path.parts) != 1
    ):
        raise ValueError(f"Invalid ticket ID '{ticket_id}'. It should be a simple name without slashes or traversal.")
    
    base_path = Path(output_arg or cfg.get("output_dir") or os.getcwd()).expanduser()
    if base_path.exists() and not base_path.is_dir():
        raise ValueError(f"Output path '{base_path}' is not a directory.")
    
    out = base_path / ticket_path.name
    try:
        out.mkdir(parents=True, exist_ok=True)
    except OSError as err:
        raise ValueError(f"Could not create output directory '{out}': {err}") from err
    
    return out

# -------------------------------
# checking select file is word or not
# -------------------------------
def is_word_file(path: str) -> bool:
    # Check if it is a file (not a folder) and ends with .docx or .doc
    return os.path.isfile(path) and path.lower().endswith((".doc", ".docx"))


# -------------------------------
# autofit row
# -------------------------------
def auto_adjust_row_height(sheet, row_num, columns, base_height=15):
    max_lines = 1

    for col in columns:
        cell = sheet[f'{col}{row_num}']
        if cell.value:
            text = str(cell.value)

            # Get column width (default ~8.43 if None)
            col_width = sheet.column_dimensions[col].width or 8.43

            # Estimate characters per line
            chars_per_line = int(col_width * 1.2)

            # Count wrapped lines
            lines = 0
            for line in text.split("\n"):
                lines += math.ceil(len(line) / chars_per_line)

            max_lines = max(max_lines, lines)

    # Set row height
    sheet.row_dimensions[row_num].height = base_height * max_lines



def calculateEffort(
    date1: str| datetime, 
    date2: str| datetime
) -> int:
    """
    Calculate efforts for Impact anaylsis sheet (per day - 6 hrs working )
    """
    date1 = parse_string_to_dateTime(date1)
    date2 = parse_string_to_dateTime(date2)
    current = date1
    days=0
    while(current <= date2):
        day = current.weekday() # Monday is 0, Sunday is 6
        if(day != 5 and day !=6):
            days += 1
        current += timedelta(days=1)
    total_hrs = days * 6
    return total_hrs

