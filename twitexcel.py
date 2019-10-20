import xlwings as xw
import os
import tweepy
import pandas as pd
import seaborn as sns
sns.set(style="ticks", color_codes=True)
import numpy as np
from datetime import timedelta
from dotenv import load_dotenv

def main():
    wb = xw.Book.caller()
    sht = wb.sheets("Sheet1")
    
    # Change working directory to read .env file properly
    path = wb.fullname
    wd = "/".join(path.replace("\\","\\\\").split('\\')[:-1])
    wd = wd.replace("\\\\","/")
    os.chdir(wd)
    
    # Load .env
    load_dotenv()
    
    # Keys
    consumer_key = os.getenv('CONSUMER_KEY')
    consumer_secret = os.getenv('CONSUMER_SECRET')
    access_token = os.getenv('ACCESS_TOKEN')
    access_token_secret = os.getenv('ACCESS_TOKEN_SECRET')
    
    # User Inputs
    user_term = sht.range('B1').value
    user_n = sht.range('B2').value

    # Scrape
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    api = tweepy.API(auth)
    results = tweepy.Cursor(api.search, q=user_term, tweet_mode='extended', lang='en').items(user_n)
    
    ROW=5
    
    # Clear Contents
    sht.range(sht.range('A'+str(ROW)), sht.range('F'+str(ROW)).end('down')).clear_contents()
        
    for tweet in results:
        sht.range('A'+str(ROW)).value = tweet.id_str
        sht.range('B'+str(ROW)).value = tweet.author._json['screen_name']
        sht.range('C'+str(ROW)).value = tweet.author._json['location']
        sht.range('D'+str(ROW)).value = tweet.created_at + timedelta(hours=8) # Set to local time
        
        if 'retweeted_status' in tweet._json:
            sht.range('E'+str(ROW)).value = tweet._json['retweeted_status']['full_text']
        else:
            sht.range('E'+str(ROW)).value = tweet.full_text

        ROW=ROW+1
    
    # Dataframe for charts
    df = sht.range('A4').expand().options(pd.DataFrame).value
    # Sentiment score placeholder
    df['Sentiment Score'] = np.random.randint(0,100,size=(len(df),1))
    # Sentiment score placeholder
    index = ['Created At','Sentiment Score','Location']
    df = df[index]
    df_week = df
    df_week['week'] = df_week['Created At'].dt.week
    
    #Box plot
    box = sns.catplot(x='week', y='Sentiment Score', hue='Location', kind="box", data=df_week)
    box.fig.set_size_inches(7, 3)
    rng = sht.range("F1")
    sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Box Plot', update = True)
    
    #Violin plot
    box = sns.catplot(x='Location', y='Sentiment Score', kind="violin", data=df)
    box.fig.set_size_inches(7, 3)
    rng = sht.range("F16")
    sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Violin Plot', update = True)
    
if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()

    