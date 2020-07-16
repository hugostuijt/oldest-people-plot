import pandas as pd
from pathlib import Path
import datetime
from dateutil.relativedelta import relativedelta
from datetime import timedelta
from matplotlib.ticker import MultipleLocator
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.dates as mdates
import statsmodels.api as sm

path = Path(r"oldest people.xlsx")

data = pd.read_excel(path)

data['Birth'] = pd.to_datetime(data['Birth'])
dates = data['Date']  # date when prev person died, and next person is the oldest.
births = data['Birth']

# Daily dates from start of data until today
dates_x = pd.date_range(start=dates[0], end=datetime.datetime.today())

loc = 0
# list with lists for x-values of all seperate plots
x_axes = [[] for i in range(0, len(dates))]
# list with lists for ages (y-values) of all seperate plots
ages = [[] for i in range(0, len(dates))]
# Loop over every day
for i in range(0, len(dates_x)):
    rd = relativedelta(dates_x[i], births[loc])
    ages[loc].append(rd.years + rd.months / 12 + rd.days / 365)  # convert to years with decimals
    x_axes[loc].append(dates_x[i])
    # oldest person died, take the next one
    if (dates_x[i].date() >= dates[min(loc + 1, len(dates) - 1)].date()) and (loc < len(dates) - 1):
        loc = loc + 1

# ====== ALL DATA AVAILABLE, NOW READY TO PLOT ====================

left = 300  # Used to add some spacing to the left
m_f = data['M/F']
fig, ax = plt.subplots()
plt.axvspan(x_axes[0][0] - timedelta(days=left), x_axes[0][0], color='#f18e89', alpha=.25,
            lw=0)  # create pink background for left most part
all = []
temp_x = []
# iterate over all persons
countMale = 0
# Loop over all persons
for i in range(0, len(ages)):
    x = x_axes[i]
    color = '#f18e89'
    if m_f[i] == 'M':
        countMale = countMale + 1
        color = '#1565ba'  # set background to blue if male
    plt.axvspan(x[0], x[-1] + timedelta(days=1), color=color, alpha=.25, lw=0)  # set background

    all.extend(ages[i])
    temp_x.extend(x)
    plt.plot(x, ages[i], '-k')  # plot each person separately


# These are the OLS values for the dashed line, I calculated it with separate software.
x_values = np.arange(0, len(dates_x), 1)
X = pd.DataFrame(x_values, columns=['time'])
X['constant'] = 1   # add constant to regression
regression = sm.OLS(all, X).fit()
c = regression.params['constant']
x_c = regression.params['time']
ls = c + x_values * x_c

# Same OLS, but now with Jeanne Calment removed from the data (so slightly smaller slope of the dashed line)
c2 = 109.7447
x_c2 = 0.000311
ls2 = c2 + x_values * x_c2


# Use the following 2 lines to add a 10 year moving average
# window = pd.Series(all).rolling(window=10*365).mean()
# plt.plot(temp_x, window, '--r', lw=1)

# plot OLS or optionally OLS with Jeanne Calment removed
plt.plot(dates_x, ls, '--k', lw=1)
# plt.plot(dates_x, ls2, '--r', lw=1)

# Set minor/major locators for axes
ax.yaxis.set_minor_locator(MultipleLocator(.5))
ax.yaxis.set_major_locator(MultipleLocator(1))

ax.xaxis.set_minor_locator(mdates.YearLocator(1))
ax.xaxis.set_major_locator(mdates.YearLocator(5))
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))

plt.xlim(dates_x[0] - timedelta(days=left), dates_x[-1])
plt.ylim(104, 124)
plt.title("Age of the oldest known living person over time")

plt.setp(ax.get_xticklabels(), rotation=30, horizontalalignment='right')
plt.xlabel('Date')
plt.ylabel('Age')
print("Percentage male: " + str(countMale / len(dates)))

plt.show()
plt.grid(alpha=.25)  # add gridlines

# Save to png
# plt.savefig(r"plot.png", bbox_inches='tight', dpi=500)
