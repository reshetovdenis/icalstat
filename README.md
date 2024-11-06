# Calendar Summary Report Generator

This Python script generates a summary report of Apple Calendar events over a specified period and exports the data into an Excel spreadsheet with charts. The report includes details of event durations per calendar, formatted and color-coded according to the calendar colors, and visualized with bar and pie charts.
## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Features

- **Access and Authorization**: Requests access to the Apple Calendar, with checks for authorization status.
- **Date Range Filtering**: Generates a report for events within a specified range of days, up to the current date.
- **Calendar Color Coding**: Extracts and uses the calendar color to visually distinguish entries in the Excel report.
- **Excel Report with Charts**:
  - Main data sheet, organized by event date and calendar name with total duration per calendar per day.
  - Stacked bar chart sheet that shows event duration by day and calendar.
  - Pie chart sheet summarizing total duration per calendar.

## Requirements

- **macOS** (uses Apple’s EventKit framework)
- **Python Packages**: `argparse`, `openpyxl`
  - Install using `pip install openpyxl`.

## Usage

### Running the Script

Run the script from the terminal with optional arguments to specify the number of days and output path:

```bash
python report.py [DAYS] [-o OUTPUT]