import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import datetime
from plotly.offline import init_notebook_mode, plot, iplot
import cufflinks as cf
#import shutil
import os
init_notebook_mode()

now = datetime.datetime.now()
# Where are CSV files?
daily = ('V:/Common Information/Digital/Reports/Error reports/source/csv/daily.csv')
hourly = ('V:/Common Information/Digital/Reports/Error reports/source/csv/hourly.csv')
errors = {row[0] : row[1] for _, row in pd.read_csv("V:/Common Information/Digital/Reports/Error reports/source/errors.csv",delimiter=';').iterrows()}
# Where to save?
report_place =str('V:/Common Information/Digital/Reports/Error reports/' + now.strftime("%Y-%m-%d"))
if not os.path.isdir(report_place):
    os.makedirs(report_place)
if not os.path.isdir(report_place + '/data/'):
    os.makedirs(report_place + '/data/')
daily_r = (report_place + '/data/'+ '/Daily.pdf')

hourly_r = (report_place +'/data/'+ '/Hourly.pdf')


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
                  title=name,
                  hoverlabel=dict(
                      bgcolor='black', ),
                  autosize = True)
    plot(data.iplot(asFigure=True, layout=layout, kind='line', title=period),
                 filename=report_place + '/data/' + name + '.html', auto_open=False)
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
        '''
        offline.plot(data[i].iplot(asFigure=True, title = i, kind='line', dimensions=(1800, 1000)),
                     filename=report_place + '/' + i + '.html', auto_open=False)
        plt.clf()
        '''
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
top = top.rename(index=str, columns={"0Number of errors": "Number of errors", "0Name of errors": "Name of errors"})
print(top)
# Daily report
period = '10 days'
name = 'Daily'
with PdfPages(daily_r) as pdf:
    main_to_pdf(data, period)
    # Total converting to a table in PDF
    print_top(top)
    err_graph(top, data)
top_daily = top
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
name = 'Hourly'
with PdfPages(hourly_r) as pdf:
    # Main graph
    main_to_pdf(data, period)
    # Total converting to a table in PDF
    print_top(top)
    # Plotting and saving top graphics into PDF
    err_graph(top, data)


# Zip resulted file
#shutil.make_archive(report_place), 'zip', reportplace)



f = open(report_place+'/Report.html','w')
files = os.listdir(report_place + '/data/')
i=0
#files ='<br>'.join(list_of_files)
while i < len(files):
    files[i]=str('<p><a href=' + 'data/' + files[i]  + '>' + files[i] + '</a></p>')
    i += 1
files =''.join(files)
message = str(("<html>\n"
           "<head></head>\n"
           "<body>" + files + "<br></body>\n"
           "</html>"))

f.write(message)
f.close()
