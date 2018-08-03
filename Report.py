import pandas as pd
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# import matplotlib as mpl
# import matplotlib.axes.Axes as ax

# Where are CSV files?
daily = ('~/Desktop/p_report/csv/daily.csv')
hourly = ('~/Desktop/p_report/csv/hourly.csv')

# Where to save?

daily_r = ('C:/Users/313457/Desktop/p_report/Daily report.pdf')
hourly_r = ('C:/Users/313457/Desktop/p_report/Hourly report.pdf')


# Total converting to a table in PDF
def print_top(top):
    total_rows, total_cols = top.shape  # There were 3 columns in my df

    rows_per_page = 40  # Assign a page cut off length
    rows_printed = 0
    page_number = 1

    while (total_rows > 0):
        # put the table on a correctly sized figure
        fig = plt.figure(figsize=(8.5, 11))
        plt.gca().axis('off')
        matplotlib_tab = pd.plotting.table(plt.gca(), top.iloc[rows_printed:rows_printed + rows_per_page],
                                           loc='upper center', colWidths=[0.2, 0.2, 0.2])

        # Give you cells some styling
        table_props = matplotlib_tab.properties()
        table_cells = table_props['child_artists']  # I have no clue why child_artists works
        for cell in table_cells:
            cell.set_height(0.024)
            cell.set_fontsize(12)

        # Add a header and footer with page number
        fig.text(4.25 / 8.5, 10.5 / 11., "Top errors", ha='center', fontsize=12)
        fig.text(4.25 / 8.5, 0.5 / 11., 'A' + str(page_number), ha='center', fontsize=12)

        pdf.savefig()
        plt.clf()

        # Update variables
        rows_printed += rows_per_page
        total_rows -= rows_per_page
        page_number += 1
    return(0)

# Plotting and saving top graphics into PDF
def err_graph(top,data):
    error = 0
    while error < 10:
        i = top.index[error]
        data[i].plot(figsize=(20, 10))
        plt.xticks(data.index.values, data['time'], rotation=60)
        plt.title(i)
        plt.grid(True)
        plt.draw()
        pdf.savefig()
        error = error + 1
        plt.clf()

    return(0)

# Daily report start
data = pd.read_csv(daily)

# Date reformatting

data['day'] = data._time.str.split('T').str.get(0)
data['hour'] = data._time.str.split(':00.000-0400').str.get(0)
data['hour'] = data.hour.str.split('T').str.get(1)
data['time'] = data['day']
data = data.drop(['_time', 'day', 'hour'], axis=1)
data = data[['time'] + data.columns[:-1].tolist()]
data.set_index('time')

# Date conversion

# data.index = pd.to_datetime(data.index)
# data.index = data.index.tz_localize(tz ='Etc/GMT+4')
# data.index = data.index.tz_convert('Etc/GMT-3')

# Daily top counting

top = data.sum(axis=0)
type(top)
top = pd.Series.to_frame(top)
top = top.drop(top.index[0])
top = top.sort_values(by=0, ascending=False)
top.index = top.index.map(str)

# Daily report

from matplotlib.backends.backend_pdf import PdfPages

#with PdfPages('Daily report.pdf') as pdf:
with PdfPages(daily_r) as pdf:
    data.plot(linestyle='-', marker='o', figsize=(30, 15))
    plt.rcParams['axes.formatter.useoffset'] = False
    plt.xticks(data.index.values, data['time'], rotation=60)
    plt.grid()
    plt.axhline(y=0, color='k')
    plt.axvline(x=0, color='k')
    plt.title('Errors during last day')
    pdf.savefig()
    plt.clf()


# Total converting to a table in PDF

    total_rows, total_cols = top.shape  # There were 3 columns in my df

    rows_per_page = 40  # Assign a page cut off length
    rows_printed = 0
    page_number = 1

    print_top(top)
    err_graph(top, data)

#Hourly report

data = pd.read_csv(hourly)

# Date reformatting

data['day'] = data._time.str.split('T').str.get(0)
data['hour'] = data._time.str.split(':00.000-0400').str.get(0)
data['hour'] = data.hour.str.split('T').str.get(1)
data['time'] = data['day'] + ', ' + data['hour']
data = data.drop(['_time', 'day', 'hour'], axis=1)
data = data[['time'] + data.columns[:-1].tolist()]
data.set_index('time')

# Daily top counting

top = data.sum(axis=0)
top = pd.Series.to_frame(top)
top = top.drop(top.index[0])
top = top.sort_values(by=0, ascending=False)
top.index = top.index.map(str)


# Hourly report

from matplotlib.backends.backend_pdf import PdfPages

#with PdfPages('Hourly report.pdf') as pdf:
with PdfPages(hourly_r) as pdf:
    data.plot(linestyle='-', marker='o', figsize=(30, 15))
    plt.rcParams['axes.formatter.useoffset'] = False
    plt.xticks(data.index.values, data['time'], rotation=60)
    plt.grid()
    plt.axhline(y=0, color='k')
    plt.axvline(x=0, color='k')
    plt.title('Errors during last day')
    pdf.savefig()
    plt.clf()


    # Total converting to a table in PDF

    print_top(top)
    # Plotting and saving top graphics into PDF
    err_graph(top,data)
