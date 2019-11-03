import xlwings as xw
import os
import tweepy
import pandas as pd
import numpy as np
import seaborn as sns
import numpy as np
from datetime import timedelta
from dotenv import load_dotenv
import win32api
import win32con
import spacy
from PIL import Image
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
    
    if user_n == 0:
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please ensure that the Maximum # of Tweets is greater than 0.', 'Exception', win32con.MB_ICONINFORMATION)
        return
    
    if retweet == None:
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'Please indicate whether retweets should be included.', 'Exception', win32con.MB_ICONINFORMATION)

    # Construct Query
    search_query = construct_query(poster, search_terms, hashtags, retweet)
    
    # Scrape
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    api = tweepy.API(auth, wait_on_rate_limit=True)
    results = tweepy.Cursor(api.search, q=search_query, tweet_mode='extended', lang='en', count=user_n).items(user_n+1)
        
    # Prompt user if no results are found
    if not(any(True for _ in results)):
        win32api.MessageBox(xw.apps.active.api.Hwnd, 'No tweets found. Please refine the search criteria.', 'Exception', win32con.MB_ICONINFORMATION)
        return
    
    # Clear Contents
    out_sht.range(out_sht.range('A2'), out_sht.range('A2').end('down')).api.EntireRow.Delete()

    results_list = []
    for tweet in results:
        tweet_info = {}
        tweet_info['Created At'] = tweet.created_at + timedelta(hours=8) # Set to local time
        tweet_info['Screen Name'] = tweet.author._json['screen_name']
        tweet_info['User Location'] = tweet.author._json['location']
        tweet_info['Followers'] = tweet.author._json['followers_count']
        tweet_info['Following'] = tweet.author._json['friends_count']
        
        if 'retweeted_status' in tweet._json:
            tweet_info['Is Retweet'] = 'Yes'
            tweet_info['Likes'] = tweet._json['retweeted_status']['favorite_count']
            tweet_text = tweet._json['retweeted_status']['full_text']
        else:
            tweet_info['Is Retweet'] = 'No'
            tweet_info['Likes'] = tweet._json['favorite_count']
            tweet_text = tweet.full_text
        
        tweet_info['Full Text'] = re.sub(r"http\S+", "", tweet_text) # Clean http links
        tweet_info['Sentiment Score'] = vader_compound_score(tweet_text)
        results_list.append(tweet_info)
    
    results_df = pd.DataFrame(results_list, columns=tweet_info.keys())
    out_sht.range('A1').options(index=False).value = results_df
    
    # Column Widths
    out_sht.range('A1').column_width = 14.88 #Created At
    out_sht.range('B1').column_width = 16.25 #Screen Name
    out_sht.range('C1').column_width = 16.5  #User Location
    out_sht.range('D1').column_width = 10.25 #Followers
    out_sht.range('E1').column_width = 10.25 #Following
    out_sht.range('F1').column_width = 11.2  #Is Retweet
    out_sht.range('G1').column_width = 7     #Likes
    out_sht.range('H1').column_width = 65    #Full Text
    out_sht.range('I1').column_width = 17.5  #Sentiment Score
    
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
    
    # create number of followers grouping
    def follower_groups (df_column):
        if (df_column <=50):
            return '0 to 50'
        elif (df_column <=100):
            return '51 to 100'
        elif (df_column <=500):
            return '101 to 500'
        elif (df_column <=1000):
            return '501 to 1000'
        elif (df_column <=5000):
            return '1001 to 5000'
        elif (df_column >5000):
            return 'More than 5000'
        return np.nan
  
    # Dataframe for charts
    #df = out_sht.range('A1').expand().options(pd.DataFrame).value
    df = results_df.copy()
    index = ['Created At','Sentiment Score','Followers','User Location', 'Full Text']
    df = df[index]
    df['Created At minute'] = df['Created At'] - pd.to_timedelta(df['Created At'].dt.second, unit='s')
    df['sentiment_score_category'] = score_groups(df['Sentiment Score'])
    df['tweet_user_followers'] = df['Followers'].apply(follower_groups)
    
    # wordcloud 
    text = ' '.join(df['Full Text'].tolist())
    stopwords = set(STOPWORDS)
    stopwords.update([r'http[s]?://(?:[a-zA-Z]|[0-9]|[[email protected]&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'])
    twitter_mask = np.array(Image.open("twitter.png"))
    twitter_mask[twitter_mask>245]= 255
    wordcloud = WordCloud(stopwords=stopwords, background_color="white", mask = twitter_mask,
                          max_words=100, contour_width=5, contour_color='lightblue').generate(text)
    plt.axis("off")
    wordcloud_fig = plt.imshow(wordcloud, interpolation='bilinear').get_figure()
    wordcloud_fig.set_size_inches(2.5, 2)
    rng = viz_sht.range("A3")
    viz_sht.pictures.add(wordcloud_fig, top=rng.top, left=rng.left, name='Word Cloud', update = True)
    
    #Horizontal bar plot by number of followers
    df_followers = df.groupby(['tweet_user_followers']).size().reset_index().rename(columns={0:'counts'})
    followers_order=['0 to 50', '51 to 100', '101 to 500', '501 to 1000', '1001 to 5000', 'More than 5000', ]
    followers_order.reverse()
    palette = sns.color_palette("Blues")
    palette.reverse()
    bar = sns.catplot(x="counts", y="tweet_user_followers", order = followers_order, kind="bar", palette = palette, data=df_followers)
    bar.set_titles("{Sentimental Score Distribution}")
    bar.fig.set_size_inches(3.5, 2.7)
    rng = viz_sht.range("E23")
    viz_sht.pictures.add(bar.fig, top=rng.top, left=rng.left, name='Bar Plot followers', update = True)
    
    #Violin plot for sentiment scores by number of followers
    followers_order=['0 to 50', '51 to 100', '101 to 500', '501 to 1000', '1001 to 5000', 'More than 5000', ]
    palette = sns.color_palette("Blues")
    bar = sns.catplot(x="tweet_user_followers", y="Sentiment Score", order = followers_order, kind="violin", palette = palette, data=df)
    bar.set_titles("{Sentimental Score Distribution}")
    bar.set_xticklabels(rotation=70)
    bar.fig.set_size_inches(3.5, 2.7)
    rng = viz_sht.range("M23")
    viz_sht.pictures.add(bar.fig, top=rng.top, left=rng.left, name='Violin Plot followers', update = True)
    
    #Bar plot by sentiment score category
    df_score = df.groupby(['sentiment_score_category']).size().reset_index().rename(columns={0:'counts'})
    score_order=['-1.0 to -0.8','-0.8 to -0.6','-0.6 to -0.4','-0.4 to -0.2','-0.2 to 0.0','0.0 to 0.2','0.2 to 0.4','0.4 to 0.6','0.6 to 0.8','0.8 to 1.0']
    palette = sns.color_palette("RdBu",10)
    bar = sns.catplot(x="sentiment_score_category", y="counts", order = score_order, kind="bar", palette = palette, data=df_score)
    bar.set_titles("{Sentimental Score Distribution}")
    bar.set_xticklabels(rotation=70)
    bar.fig.set_size_inches(3.5, 2.7)
    rng = viz_sht.range("E5")
    viz_sht.pictures.add(bar.fig, top=rng.top, left=rng.left, name='Bar Plot score', update = True)
    
    #Line plot by sentiment score
    line = sns.catplot(x="Created At", y="Sentiment Score", kind="point", data=df)
    line.set_titles("{Sentimental Score Timeline}")
    line.set(xticks=[],xlabel='')
    line.fig.set_size_inches(7, 2.7)
    rng = viz_sht.range("M5")
    viz_sht.pictures.add(line.fig, top=rng.top, left=rng.left, name='Line Plot score', update = True)
    
    #Box plot
#    box = sns.catplot(x='week', y='Sentiment Score', hue='Location', kind="box", data=df)
#    box.fig.set_size_inches(7, 3)
#    rng = out_sht.range("G16")
#    out_sht.pictures.add(box.fig, top=rng.top, left=rng.left, name='Box Plot', update = True)

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
    
    
    
    
    

    