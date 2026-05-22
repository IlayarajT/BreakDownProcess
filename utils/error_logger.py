import os
import time
from bs4 import BeautifulSoup

def log_error(input_file, start_time, end_time, error_log_file):
    error_details = {
        "date": time.strftime('%Y-%m-%d'),
        "file_name": os.path.basename(input_file),
        "start_time": time.strftime('%H:%M:%S', time.localtime(start_time)),
        "end_time": time.strftime('%H:%M:%S', time.localtime(end_time))
    }
    create_error_log(error_log_file, error_details)

def create_error_log(file_name, error_details):
    if os.path.exists(file_name):
        with open(file_name, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f, 'html.parser')
            table = soup.find('table')
            if table:
                rows = table.find_all('tr')
                last_index = 0
                if rows:
                    first_td = rows[-1].find('td')
                    if first_td and first_td.text:
                        try:
                            last_index = int(first_td.text.strip())
                        except (ValueError, TypeError):
                            # First <td> is not a number — count data rows instead
                            last_index = len(rows) - 1  # subtract header row
            else:
                table = _new_table()
                soup.body.append(table)
                last_index = 0
    else:
        soup = _new_soup()
        table = soup.find('table')
        last_index = 0

    new_index = last_index + 1
    row = soup.new_tag('tr')
    # Add S.No. column first
    sno_td = BeautifulSoup(f'<td>{new_index}</td>', 'html.parser')
    row.append(sno_td)
    # Add the rest of the columns
    for key in ["date", "file_name", "start_time", "end_time"]:
        row.append(BeautifulSoup(f'<td>{error_details[key]}</td>', 'html.parser'))
    table.find('tbody').append(row)

    with open(file_name, 'w', encoding='utf-8') as f:
        f.write(soup.prettify())

def _new_soup():
    return BeautifulSoup(
        '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
        '<title>Error Log</title><link rel="stylesheet" type="text/css" href="DataTable/datatables.css">'
        '<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>'
        '<script src="DataTable/datatables.js"></script>'
        '<script>$(document).ready(function() { $("#dataTable").DataTable(); });</script>'
        '</head><body><table id="dataTable"><thead><tr><th>S. No.</th>'
        '<th>Date</th><th>File Name/Package Name</th>'
        '<th>Start Time</th><th>End Time</th></tr></thead><tbody></tbody></table></body></html>',
        'html.parser'
    )

def _new_table():
    return BeautifulSoup(
        '<table id="dataTable"><thead><tr><th>S. No.</th>'
        '<th>Date</th><th>File Name/Package Name</th>'
        '<th>Start Time</th><th>End Time</th></tr></thead><tbody></tbody></table>',
        'html.parser'
    ).table
