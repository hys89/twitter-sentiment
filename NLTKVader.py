
# # Using NLTK Vader to analyse the sentiment of tweets 
# 
# VADER (Valence Aware Dictionary and sEntiment Reasoner) is a lexicon and rule-based sentiment analysis tool that is specifically attuned to sentiments expressed in social media. 
# VADER uses a combination of A sentiment lexicon is a list of lexical features (e.g., words) which are generally labelled according to their semantic orientation as either positive or negative.
# 
# VADER has been found to be quite successful when dealing with social media texts, NY Times editorials, movie reviews, and product reviews. 
# This is because VADER not only tells about the Positivity and Negativity score but also tells us about how positive or negative a sentiment is.
# 
# VADER has a lot of advantages over traditional methods of Sentiment Analysis, including:
# - It works exceedingly well on social media type text, yet readily generalizes to multiple domains
# - It doesn’t require any training data but is constructed from a generalizable, valence-based, human-curated gold standard sentiment lexicon (Amazon Mechanical Turk labelling)
# - It is fast enough to be used online with streaming data, and
# - It does not severely suffer from a speed-performance tradeoff.
# 
# To install Vader, use 
# > pip install vaderSentiment 
# You may also need to install 'requests' library too. 
# 
# More info: https://medium.com/analytics-vidhya/simplifying-social-media-sentiment-analysis-using-vader-in-python-f9e6ec6fc52f

from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
analyser = SentimentIntensityAnalyzer()

def vader_compound_score(sentence):
    score = analyser.polarity_scores(str(sentence))
    # print("{:-<40} {}".format(sentence, str(score)))
    return score['compound']


if __name__ == "__main__":
    
    print('Testing.. You should see some example tweets and their scores (from -1 to 1)')
    print('\n')

    vader_compound_score("Four people injured after reports of stabbings at the Arndale shopping centre in Manchester")
    print('\n')

    vader_compound_score("Turkish troops enter northern Syria, says President Erdogan, setting up a potential clash with Kurdish-led forces")
    print('\n')

    vader_compound_score("Always remember Hong Kong and all of its beauty")
    print('\n')

    vader_compound_score(
                        """Breaking News & Video! Blizzard Makes Things EVEN WORSE!  
    Just As The #BoycottBlizzard Noise Was Starting To Calm Down They Retro-Actively BANNED The American Players Who Showed Support For Hong Kong! 
    They Are Trying To Suppress This! Watch & Share Far!""")
    print('\n')

    vader_compound_score("Kenyan Eliud Kipchoge becomes the first athlete to run a marathon in under two hours, completing the 26.2 miles in 1:59:40")