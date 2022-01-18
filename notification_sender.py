import time
from plyer import notification
import pandas as pd
from random import randrange


while True:
    time.sleep(5)
    try:
        df = pd.read_excel('wordlist.xlsx')
        df.fillna('Not Available')
    except:
        pass

    # Getting number of row and selecting a random row
    row = df.shape[0]
    rand_row = randrange(0, row)

    # Send desktop notification
    notification.notify(
        title = str(df.iloc[rand_row, 0]) + ': ' + str(df.iloc[rand_row, 1]),
        message = str(df.iloc[rand_row, 2]),
        app_name = 'Word Reminder',
        app_icon = 'dictionary.ico',
        timeout = 100,
        ticker='Hello'
    )
