import fnmatch
import os

import pandas as pd
import numpy as np

from jinja2 import Environment, FileSystemLoader
from xhtml2pdf import pisa

file_format = 'toptracker*.csv'
input_file_name = ""
output_file_name = "invoice_report.pdf"

for file in os.listdir('.'):
    if fnmatch.fnmatch(file, file_format):
        input_file_name = file

if input_file_name is "":
    print("No file found with the format: {}".format(file_format))
    exit(-1)

df = pd.read_csv(input_file_name)
df['start_time'] = pd.to_datetime(df['start_time']).dt.strftime('%d-%B-%Y')
df['duration_hours'] = df['duration_seconds'].apply(lambda x: float(x) / 3600)
df['hourly_rate'] = 35.0
df['total'] = df['duration_hours'] * df['hourly_rate']
table = pd.pivot_table(df, index=["project", "start_time"], values=['duration_hours', 'hourly_rate', 'total'],
                       aggfunc=np.sum, margins=True)
print(table)

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("invoice_report.html")

template_vars = {"title": "Sales Funnel Report - National",
                 "invoice_data": table.to_html()}
html_out = template.render(template_vars)

print(html_out)

pisa.showLogging()
# open output file for writing (truncated binary)
if os.path.isfile(output_file_name):
    os.remove(output_file_name)

resultFile = open(output_file_name, "w+b")

# convert HTML to PDF
pisaStatus = pisa.CreatePDF(
    html_out,  # the HTML to convert
    dest=resultFile)  # file handle to receive result

# close output file
resultFile.close()  # close output file
