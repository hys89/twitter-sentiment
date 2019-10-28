#More info on FastText here: https://arxiv.org/abs/1607.01759 

from keras.models import load_model
from keras.datasets import imdb
from keras.preprocessing import sequence
import spacy 
import numpy as np

nlp = spacy.load('en_core_web_sm')

# Set parameters:
ngram_range = 1
max_features = 80000
maxlen = 150
batch_size = 32
embedding_dims = 50
epochs = 5

# load model
model = load_model('FastText_Unigram_150len80kfeatures.h5')
# summarize model.
model.summary()

word_index = imdb.get_word_index()

def sentence_to_indices(sentence):
    nlp_doc = nlp(sentence)
    indices_list = []
    for token in nlp_doc:
        token_text = token.text.lower()
        indices_list.append(word_index[token_text] if token_text in word_index else 0)
    sequence.pad_sequences([indices_list], maxlen = maxlen, padding = 'post')
    return(indices_list)

def fasttext_sentiment(text):
    indices = sentence_to_indices(text)
    padded_indices = sequence.pad_sequences([indices], maxlen = maxlen, padding = 'post')
    prediction = model.predict(padded_indices)
    score = (prediction - 0.8) * 2
    return round(float(score), 2) 


print('Testing.. You should see some example tweets and their scores (from -1 to 1)')

print(fasttext_sentiment("Four people injured after reports of stabbings at the Arndale shopping centre in Manchester"))

print(fasttext_sentiment("Turkish troops enter northern Syria, says President Erdogan, setting up a potential clash with Kurdish-led forces"))

print(fasttext_sentiment("Always remember Hong Kong and all of its beauty"))

print(fasttext_sentiment("""this food tastes like garbage. I am having a bad terrible day. My leg is broken and everything hurts!"""))

print(fasttext_sentiment("Kenyan Eliud Kipchoge becomes the first athlete to run a marathon in under two hours, completing the 26.2 miles in 1:59:40"))