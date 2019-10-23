import xlwings as xw
import os
import tweepy
import pandas as pd
import seaborn as sns
import numpy as np
from datetime import timedelta
from dotenv import load_dotenv
import win32api
import win32con
from NLTKVader import vader_compound_score
sns.set(style="ticks", color_codes=True)

def construct_query(poster, search_terms, hashtags):
    
    search_query = search_terms

    if poster != None and search_query != None:
    	search_query = str(search_query) + ' from:' + poster
        
    if poster != None and search_query == None:
    	search_query = 'from:' + poster
        
    if hashtags != None and search_query != None:
        hashtag_string = ['#'+h for h in hashtags.split()]
        search_query = str(search_query) + ' ' + ' '.join(hashtag_string)
        
    if hashtags != None and search_query == None:
        hashtag_string = ['#'+h for h in hashtags.split()]
        search_query = ' '.join(hashtag_string)
        
    return search_query

def cell_color(sentiment):
    if -1 <= sentiment < -0.8:
        color = (165,0,38)
    elif -0.8 <= sentiment < -0.6:
        color = (215,48,39)
    elif -0.6 <= sentiment < -0.4:
        color = (244,109,67)
    elif -0.4 <= sentiment < -0.2:
        color = (253,174,97)
    elif -0.2 <= sentiment < 0:
        color = (254,224,139)
    elif sentiment == 0:
        color = (255,255,191)
    elif 0 < sentiment <= 0.2:
        color = (217,239,139)
    elif 0.2 < sentiment <= 0.4:
        color = (166,217,106)
    elif 0.4 < sentiment <= 0.6:
        color = (102,189,99)
    elif 0.6 < sentiment <= 0.8:
        color = (26,152,80)
    elif 0.8 < sentiment <= 1:
        color = (0,104,55)
    return color
    

def main():
    wb = xw.Book.caller()
    in_sht = wb.sheets("User Interface")
    out_sht = wb.sheets("Results")
    
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
    search_terms = in_sht.range('B10').value
    hashtags = in_sht.range('B11').value
    poster = in_sht.range('B12').value
    user_n = in_sht.range('B13').value
    
    # Check for exceptions
    if search_terms == None and poster == None and hashtags == None:
    	win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please key in at least one of the following: Search Term, Hashtag, Poster', 'Exception', win32con.MB_ICONINFORMATION)
    	return
    
    if ((poster != None) and (len(poster.split()) > 1)):
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'Multiple Posters detected. Please key in only one Poster.', 'Exception', win32con.MB_ICONINFORMATION)
        return

    if user_n == None:
    	win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please key in the Maximum # of Tweets to be returned', 'Exception', win32con.MB_ICONINFORMATION)
    	return
    
    # Construct Query
    search_query = construct_query(poster, search_terms, hashtags)
    
    # Scrape
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    api = tweepy.API(auth)
    results = tweepy.Cursor(api.search, q=search_query, tweet_mode='extended', lang='en').items(user_n+1)

    # Prompt user if no results are found
    if not(any(True for _ in results)):
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'No tweets found. Please refine the search criteria.', 'Exception', win32con.MB_ICONINFORMATION)
        return
    
    ROW=2
    
    # Clear Contents
    out_sht.range(out_sht.range('A'+str(ROW)), out_sht.range('F'+str(ROW)).end('down')).clear_contents()
    out_sht.range(out_sht.range('F'+str(ROW)), out_sht.range('F'+str(ROW)).end('down')).color = None
        
    for tweet in results:
        out_sht.range('A'+str(ROW)).value = tweet.id_str
        out_sht.range('B'+str(ROW)).value = tweet.author._json['screen_name']
        out_sht.range('C'+str(ROW)).value = tweet.author._json['location']
        out_sht.range('D'+str(ROW)).value = tweet.created_at + timedelta(hours=8) # Set to local time
        
        if 'retweeted_status' in tweet._json:
            tweet_text = tweet._json['retweeted_status']['full_text']
        else:
            tweet_text = tweet.full_text
        out_sht.range('E'+str(ROW)).value = tweet_text
        
        vader_score = vader_compound_score(tweet_text)
        out_sht.range('F'+str(ROW)).value = vader_score
        out_sht.range('F'+str(ROW)).color = cell_color(vader_score)

        ROW=ROW+1

    #######################################
    ### Charts ############################
    #######################################

    # Dataframe for charts
    df = out_sht.range('A1').expand().options(pd.DataFrame).value
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
    rng = out_sht.range("G1")
    out_sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Box Plot', update = True)
    
    #Violin plot
    box = sns.catplot(x='Location', y='Sentiment Score', kind="violin", data=df)
    box.fig.set_size_inches(7, 3)
    rng = out_sht.range("G16")
    out_sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Violin Plot', update = True)


    #######################################
    ### Prompt Completion #################
    #######################################

    win32api.MessageBox(xw.apps.active.api.Hwnd, 'The sentiment analysis is complete! Check out the Results sheet!', 'Done', win32con.MB_ICONINFORMATION)

    
if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()

    