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
import spacy
import re
from NLTKVader import vader_compound_score
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
sns.set(style="ticks", color_codes=True)

def construct_query(poster, search_terms, hashtags, retweet):
    
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
    
    if retweet == "No":
        search_query = search_query + " -filter:retweets"
        
    return search_query    

def main(dest):
    wb = xw.Book.caller()
    in_sht = wb.sheets("User Interface")
    out_sht = wb.sheets("Tweets")
    viz_sht = wb.sheets("Dashboard")
    
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
    retweet = in_sht.range('B14').value
    
    # Check for exceptions
    if search_terms == None and poster == None and hashtags == None:
    	win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please key in at least one of the following: Search Term, Hashtag, Poster', 'Exception', win32con.MB_ICONINFORMATION)
    	return
    
    if ((poster != None) and (len(poster.split()) > 1)):
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'Multiple Posters detected. Please key in only one Poster.', 'Exception', win32con.MB_ICONINFORMATION)
        return

    if user_n == None:
    	win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please key in the Maximum # of Tweets to be returned.', 'Exception', win32con.MB_ICONINFORMATION)
    	return
    
    if retweet == None:
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please indicate whether retweets should be included.', 'Exception', win32con.MB_ICONINFORMATION)

    # Construct Query
    search_query = construct_query(poster, search_terms, hashtags, retweet)
    
    # Scrape
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    api = tweepy.API(auth)
    results = tweepy.Cursor(api.search, q=search_query, tweet_mode='extended', lang='en').items(user_n+1)
        
    # Prompt user if no results are found
    if not(any(True for _ in results)):
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'No tweets found. Please refine the search criteria.', 'Exception', win32con.MB_ICONINFORMATION)
        return
    
    # Clear Contents
    out_sht.range(out_sht.range('A2'), out_sht.range('H2').end('down')).clear_contents()

    results_list = []
    for tweet in results:
        tweet_info = {}
        tweet_info['ID'] = tweet.id_str
        tweet_info['Created At'] = tweet.created_at + timedelta(hours=8) # Set to local time
        tweet_info['Screen Name'] = tweet.author._json['screen_name']
        tweet_info['User Location'] = tweet.author._json['location']
        tweet_info['Followers'] = tweet.author._json['followers_count']
        tweet_info['Following'] = tweet.author._json['friends_count']
        
        if 'retweeted_status' in tweet._json:
            tweet_text = tweet._json['retweeted_status']['full_text']
        else:
            tweet_text = tweet.full_text
        
        tweet_info['Full Text'] = re.sub(r"http\S+", "", tweet_text) # Clean http links
        tweet_info['VADER Sentiment'] = vader_compound_score(tweet_text)
        results_list.append(tweet_info)
    
    results_df = pd.DataFrame(results_list, columns=tweet_info.keys())
    out_sht.range('A1').options(index=False).value = results_df
    
    # Column Widths
    out_sht.range('A1').column_width = 11.25 #ID
    out_sht.range('B1').column_width = 14.88 #Created At
    out_sht.range('C1').column_width = 16.25 #Screen Name
    out_sht.range('D1').column_width = 25    #User Location
    out_sht.range('E1').column_width = 10.25 #Followers
    out_sht.range('F1').column_width = 10.25 #Following
    out_sht.range('G1').column_width = 83    #Full Text
    out_sht.range('H1').column_width = 17.5  #VADER
    
    # Autofit Rows
    out_sht.autofit("rows")
    
    # Bold Headers
    out_sht.range('1:1').api.Font.Bold = True

    #######################################
    ### Charts ############################
    #######################################

    # create sentiment score grouping by interval
    def score_groups (df_column, interval = 0.2):
        lower = df_column.add(0.0001).mul(1/interval).apply(np.floor)*interval
        upper = df_column.add(0.0001).mul(1/interval).apply(np.ceil)*interval
        category = lower.round(1).astype(str) + ' to ' + upper.round(1).astype(str)
        return (category)
    
    # Dataframe for charts
    df = out_sht.range('A1').expand().options(pd.DataFrame).value
    index = ['Created At','VADER Sentiment','User Location', 'Full Text']
    df = df[index]
    df['week'] = df['Created At'].dt.week
    df['score_category'] = score_groups(df['VADER Sentiment'])
    
    # wordcloud 
    text = ' '.join(df['Full Text'].tolist())
    stopwords = set(STOPWORDS)
    stopwords.update([r'http\S+'])
    wordcloud = WordCloud(stopwords=stopwords, background_color="white").generate(text)
    plt.axis("off")
    wordcloud_fig = plt.imshow(wordcloud, interpolation='bilinear').get_figure()
    rng = viz_sht.range("a1")
    viz_sht.pictures.add(wordcloud_fig, top=rng.top, left=rng.left, name='Word Cloud', update = True)
    
    #Bar plot
    df_score = df.groupby(['score_category']).size().reset_index().rename(columns={0:'counts'})
    order=['-1.0 to -0.8','-0.8 to -0.6','-0.6 to -0.4','-0.4 to -0.2','-0.2 to 0.0','0.0 to 0.2','0.2 to 0.4','0.4 to 0.6','0.6 to 0.8','0.8 to 1.0']
    bar = sns.catplot(x="score_category", y="counts", order = order, hue="score_category", kind="bar", palette = "RdBu", data=df_score)
    bar.set_xticklabels(rotation=30)
    bar.fig.set_size_inches(7, 3)
    rng = viz_sht.range("A33")
    viz_sht.pictures.add(bar.fig, top=rng.top, left=rng.left, name='Bar Plot', update = True)
    
    #Box plot
#    box = sns.catplot(x='week', y='VADER Sentiment', hue='Location', kind="box", data=df)
#    box.fig.set_size_inches(7, 3)
#    rng = out_sht.range("G16")
#    out_sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Box Plot', update = True)
#    
#    #Violin plot
#    box = sns.catplot(x='Location', y='VADER Sentiment', kind="violin", data=df)
#    box.fig.set_size_inches(7, 3)
#    rng = out_sht.range("G31")
#    out_sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Violin Plot', update = True)


    #######################################
    ### Prompt Completion #################
    #######################################
    
    if dest == "tweets":
        out_sht.activate()
    elif dest == "dashboard":
        viz_sht.activate()
        
    win32api.MessageBox(xw.apps.active.api.Hwnd, 'The sentiment analysis is complete!', 'Done', win32con.MB_ICONINFORMATION)

    
if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()
    
    
    
    
    

    