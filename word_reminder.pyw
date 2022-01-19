
import time
from plyer import notification
import pandas as pd
from random import randrange

if __name__ == '__main__':
    while True:  
        try:
            df = pd.read_excel('wordlist.xlsx')
            df = df.fillna('Not Available')
        except:
            time.sleep(5)
            continue

        # Getting number of row and selecting a random row
        row = df.shape[0]
        rand_row = randrange(0, row)

        # Send desktop notification
        notification.notify(
            title = str(df.iloc[rand_row, 0]) + ': ' + str(df.iloc[rand_row, 1]),
            message = str(df.iloc[rand_row, 2]),
            app_name = 'Word Reminder',
            app_icon = 'dictionary.ico',
            timeout = 10,
            ticker='Hello'
        )
        time.sleep(5*60)
        del df
