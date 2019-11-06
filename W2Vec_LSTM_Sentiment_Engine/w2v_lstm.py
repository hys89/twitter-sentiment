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

def build_w2v_lstm_and_tokenizer():

    with open('W2Vec_LSTM_Sentiment_Engine/results/tokenizer.pkl', 'rb') as handle:
        tokenizer = pickle.load(handle)

    model = load_model('W2Vec_LSTM_Sentiment_Engine/results/model.h5')

    return tokenizer, model
def decode_sentiment(score, include_neutral=True):
    SENTIMENT_THRESHOLDS = (0.4,0.7)
    if include_neutral:        
        label = 'NEUTRAL'
        if score <= SENTIMENT_THRESHOLDS[0]:
            label = 'NEGATIVE'
        elif score >= SENTIMENT_THRESHOLDS[1]:
            label = 'POSITIVE'

        return label
    else:
        return 'NEGATIVE' if score < 0.5 else 'POSITIVE'


def predict(tokenizer, model, text, SEQUENCE_LENGTH = 300, include_neutral=True):
    start_at = time.time()
    # Tokenize text
    x_test = pad_sequences(tokenizer.texts_to_sequences([text]), maxlen=SEQUENCE_LENGTH)
    # Predict
    score = model.predict([x_test])[0]
    # Decode sentiment
    label = decode_sentiment(score, include_neutral=include_neutral)

    return {"label": label, "score": float(score),
       "elapsed_time": time.time()-start_at} 
    


if __name__ == "__main__":
    tokenizer, model = build_w2v_lstm_and_tokenizer()    

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
        print('w2v_lstm predict: ' ,predict(tokenizer, model, text))
        print('-'*88)