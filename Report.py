import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import datetime
import plotly.offline as offline
#import plotly.graph_objs as go
#from plotly.offline.offline import _plot_html
import cufflinks as cf
import shutil
import os
import win32com.client as win32

now = datetime.datetime.now()

# Where are CSV files?
daily = ('~/Desktop/p_report/csv/daily.csv')
hourly = ('~/Desktop/p_report/csv/hourly.csv')
errors = pd.read_csv('~/Desktop/p_report/errors.csv', header=0, squeeze=False, delimiter=';').to_dict('series')
# Where to save?
if not os.path.isdir('C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d")):
    os.makedirs('C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d"))

daily_r = ('C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d") + '/Daily report ' + now.strftime(
    "%Y-%m-%d") + '.pdf')
hourly_r = ('C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d") + '/Hourly report ' + now.strftime(
    "%Y-%m-%d") + '.pdf')


# Main graph
def main_to_pdf(data, period):
    data.plot(linestyle='-', marker='o', figsize=(30, 15))
    plt.rcParams['axes.formatter.useoffset'] = False
    plt.xticks(range(len(data.index)), data.index, rotation=60)
    plt.grid()
    plt.axhline(y=0, color='k')
    plt.axvline(x=0, color='k')
    plt.title('Errors during last ' + period)
    plt.suptitle(now.strftime("%Y-%m-%d %H:%M"))
    pdf.savefig()
    layout = dict(hovermode='closest',
                  title=period,
                  hoverlabel=dict(
                      bgcolor='black', ),
                  )
    offline.plot(data.iplot(asFigure=True, layout=layout, kind='line', title=period, dimensions=(1800, 1000)),
                 filename='C:/Users/313457/Desktop/p_report/' + now.strftime(
                     "%Y-%m-%d") + '/ Main ' + period + ' ' + now.strftime("%Y-%m-%d") + '.html', auto_open=False)
    plt.clf()
    return (0)


# Total converting to a table in PDF
def print_top(top):
    total_rows, total_cols = top.shape  # There were 3 columns in my df

    rows_per_page = 60  # Assign a page cut off length
    rows_printed = 0
    page_number = 1

    while (total_rows > 0):
        # put the table on a correctly sized figure
        fig = plt.figure(figsize=(20, 40))
        plt.gca().axis('off')
        matplotlib_tab = pd.plotting.table(plt.gca(), top.iloc[rows_printed:rows_printed + rows_per_page],
                                           loc='upper center', colWidths=None)

        # Give you cells some styling
        table_props = matplotlib_tab.properties()
        table_cells = table_props['child_artists']  # I have no clue why child_artists works
        for cell in table_cells:
            cell.set_height(0.02)
            cell.set_fontsize(20)

        # Add a header and footer with page number
        fig.text(4.25 / 8.5, 10.5 / 11., "Top errors", ha='center', fontsize=16)
        fig.text(4.25 / 8.5, 0.5 / 11., 'A' + str(page_number), ha='center', fontsize=12)

        pdf.savefig()
        plt.clf()

        # Update variables
        rows_printed += rows_per_page
        total_rows -= rows_per_page
        page_number += 1
    return (0)


# Plotting and saving top graphics into PDF
def err_graph(top, data):
    error = 0
    while error < 10:
        i = top.index[error]
        data[i].plot(figsize=(20, 15))
        plt.xticks(range(len(data.index)), data.index, rotation=60)
        plt.title(errors[i])
        plt.grid(True)
        plt.draw()
        pdf.savefig()
        error = error + 1
        plt.clf()

    return (0)


# Daily report start
data = pd.read_csv(daily)

# Date reformatting
data['day'] = data._time.str.split('T').str.get(0)
data['hour'] = data._time.str.split(':00.000-0400').str.get(0)
data['hour'] = data.hour.str.split('T').str.get(1)
data['time'] = data['day']
data = data.drop(['_time', 'day', 'hour'], axis=1)
data = data[['time'] + data.columns[:-1].tolist()]
data.set_index('time', inplace=True)

# Date conversion
# data.index = pd.to_datetime(data.index)
# data.index = data.index.tz_localize(tz ='Etc/GMT+4')
# data.index = data.index.tz_convert('Etc/GMT-3')
# data.index = data.index.astype('str')


# Daily top counting
top = data.sum(axis=0)
type(top)
top = pd.Series.to_frame(top)
top = top.drop(top.index[0])
top = top.sort_values(by=0, ascending=False)
top.index = top.index.map(str)
errors_df = pd.DataFrame.from_dict(errors, orient='index')
top = top.join(errors_df, how='left', lsuffix='Number of errors', rsuffix='Name of errors')

# Daily report
period = '10 days'

with PdfPages(daily_r) as pdf:
    main_to_pdf(data, period)
    # Total converting to a table in PDF
    print_top(top)
    err_graph(top, data)

# Hourly report
data = pd.read_csv(hourly)

# Date reformatting
data['day'] = data._time.str.split('T').str.get(0)
data['hour'] = data._time.str.split(':00.000-0400').str.get(0)
data['hour'] = data.hour.str.split('T').str.get(1)
data['time'] = data['day'] + ', ' + data['hour']
data = data.drop(['_time', 'day', 'hour'], axis=1)
data = data[['time'] + data.columns[:-1].tolist()]
data.set_index('time', inplace=True)

# Date conversion
data.index = pd.to_datetime(data.index)
data.index = data.index.tz_localize(tz='Etc/GMT+4')
data.index = data.index.tz_convert('Etc/GMT-3')
data.index = data.index.astype('str')
data.index = data.index.map(lambda x: str(x)[:-9])

# Hourly top counting
top = data.sum(axis=0)
top = pd.Series.to_frame(top)
top = top.drop(top.index[0])
top = top.sort_values(by=0, ascending=False)
top.index = top.index.map(str)
errors_df = pd.DataFrame.from_dict(errors, orient='index')
top = top.join(errors_df, how='left', lsuffix='Number of errors', rsuffix='Name of errors')

# Hourly report
period = '35 hours'
with PdfPages(hourly_r) as pdf:
    # Main graph
    main_to_pdf(data, period)
    # Total converting to a table in PDF
    print_top(top)
    # Plotting and saving top graphics into PDF
    err_graph(top, data)


# Zip resulted file
shutil.make_archive('C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d"), 'zip',
                    'C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d"))


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Andrey.Kulagin@westernunion.ru; Vadim.Grekov@westernunion.ru '
mail.Subject = 'Error Report'
mail.Body = 'Daily and hourly error report'
#mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

# To attach a file to the email (optional):
attachment  = 'C:/Users/313457/Desktop/p_report/' + now.strftime("%Y-%m-%d") +'.zip'
mail.Attachments.Add(attachment)

mail.Send()