'''
Code adapted from https://www.kaggle.com/kernels/svzip/notebook
'''

# Utility
import re
import numpy as np
import os
from collections import Counter
import logging
import time
import pickle
import itertools


# DataFrame
import pandas as pd

# Matplot
import matplotlib.pyplot as plt

# Scikit-learn
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import confusion_matrix, classification_report, accuracy_score
from sklearn.manifold import TSNE
from sklearn.feature_extraction.text import TfidfVectorizer

# Keras
import tensorflow as tf
from tensorflow.keras.preprocessing.text import Tokenizer
from tensorflow.keras.preprocessing.sequence import pad_sequences
from tensorflow.keras.models import Sequential, load_model
from tensorflow.keras.layers import Activation, Dense, Dropout, Embedding, Flatten, Conv1D, MaxPooling1D, LSTM
from tensorflow.keras import utils
from tensorflow.keras.callbacks import ReduceLROnPlateau, EarlyStopping

# nltk
import nltk
from nltk.corpus import stopwords
from  nltk.stem import SnowballStemmer

# Word2vec
import gensim


# Set log
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.INFO)

nltk.download('stopwords')

# DATASET
DATASET_COLUMNS = ["target", "ids", "date", "flag", "user", "text"]
DATASET_ENCODING = "ISO-8859-1"
TRAIN_SIZE = 0.8

# TEXT CLENAING
TEXT_CLEANING_RE = "@\S+|https?:\S+|http?:\S|[^A-Za-z0-9]+"

# WORD2VEC 
W2V_SIZE = 300
W2V_WINDOW = 7
W2V_EPOCH = 32
W2V_MIN_COUNT = 10

# KERAS
SEQUENCE_LENGTH = 300
EPOCHS = 8
BATCH_SIZE = 1024

# SENTIMENT
POSITIVE = "POSITIVE"
NEGATIVE = "NEGATIVE"
NEUTRAL = "NEUTRAL"
SENTIMENT_THRESHOLDS = (0.4, 0.7)

# EXPORT
KERAS_MODEL = "model.h5"
WORD2VEC_MODEL = "model.w2v"
TOKENIZER_MODEL = "tokenizer.pkl"
ENCODER_MODEL = "encoder.pkl"

with open('W2Vec_LSTM_Sentiment_Engine/results/tokenizer.pkl', 'rb') as handle:
    tokenizer = pickle.load(handle)

print(tokenizer)
model = load_model('W2Vec_LSTM_Sentiment_Engine/results/model.h5')
print(model.summary())
def decode_sentiment(score, include_neutral=True):
    if include_neutral:        
        label = NEUTRAL
        if score <= SENTIMENT_THRESHOLDS[0]:
            label = NEGATIVE
        elif score >= SENTIMENT_THRESHOLDS[1]:
            label = POSITIVE

        return label
    else:
        return NEGATIVE if score < 0.5 else POSITIVE


def predict(text, include_neutral=True):
    start_at = time.time()
    # Tokenize text
    x_test = pad_sequences(tokenizer.texts_to_sequences([text]), maxlen=SEQUENCE_LENGTH)
    # Predict
    score = model.predict([x_test])[0]
    # Decode sentiment
    label = decode_sentiment(score, include_neutral=include_neutral)

    return {"label": label, "score": float(score),
       "elapsed_time": time.time()-start_at}  

print('Testing w2v LSTM')
test_texts = ["Four people injured after reports of stabbings at the Arndale shopping centre in Manchester",
"Turkish troops enter northern Syria, says President Erdogan, setting up a potential clash with Kurdish-led forces",
"Always remember Hong Kong and all of its beauty",
                    """Breaking News & Video! Blizzard Makes Things EVEN WORSE!
                    Just As The #BoycottBlizzard Noise Was Starting To Calm Down They Retro-Actively BANNED The American Players Who Showed Support For Hong Kong! 
They Are Trying To Suppress This! Watch & Share Far!""", 
"Kenyan Eliud Kipchoge becomes the first athlete to run a marathon in under two hours, completing the 26.2 miles in 1:59:40"]

# print(predict("I love the music"))
# print(predict("I hate the rain"))
# print(predict("i don't know what i'm doing"))

for text in test_texts:
    print(text)
    print('w2v_lstm predict: ' ,predict(text))
    print('-'*88)