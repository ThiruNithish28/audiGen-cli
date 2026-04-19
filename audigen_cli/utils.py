
import math
from datetime import datetime, timedelta

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



def calculateEffort(date1,date2):
  date1 = datetime.strptime(date1, '%d-%m-%Y')
  date2 = datetime.strptime(date2, '%d-%m-%Y')
  current = date1
  days=0
  while(current <= date2):
      day = current.weekday() # Monday is 0, Sunday is 6
      if(day != 5 and day !=6):
        days += 1
      current += timedelta(days=1)
  total_hrs = days * 6
  return total_hrs

