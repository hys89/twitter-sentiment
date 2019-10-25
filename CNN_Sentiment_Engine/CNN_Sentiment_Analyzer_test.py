from CNN_Sentiment_Analyzer import predict_sentiment
from CNN_Sentiment_Analyzer import print_prediction

print('The scores range from 0% to 100%, with 0% the most negative and 100% most positive')

sample_tweet = "Turkish troops enter northern Syria, says President Erdogan, setting up a potential clash with Kurdish-led forces"
print_prediction(sample_tweet)

sample_tweet = "This is bad!"
print_prediction(sample_tweet)

sample_tweet = "This is good!"
print_prediction(sample_tweet)

sample_tweet = "I am neutral"
print_prediction(sample_tweet)