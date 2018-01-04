# coding: utf-8

# mainly forking from notebook
# https://www.kaggle.com/johnfarrell/simple-rnn-with-keras-script

# ADDED
# 5x scaled test set
# category name embedding
# some small changes like lr, decay, batch_size~

import os
import gc
import time
start_time = time.time()
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import scipy
from sklearn.linear_model import Ridge, LogisticRegression
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.preprocessing import LabelBinarizer, LabelEncoder
from scipy.sparse import csr_matrix, hstack


import docx2txt
import os
import comtypes.client
import PyPDF2
import pandas as pd

def convert_doc2docx(full_file_name):
    in_file = full_file_name
    out_file = full_file_name.replace(".doc", ".docx")  # name of output file added to the current working directory
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)  # name of input file
    doc.SaveAs(out_file, FileFormat=16)  # output file format to Office word Xml default (code=16)
    doc.Close()
    word.Quit()
    return out_file

def read_docx(full_file_name):

    fullText = docx2txt.process(full_file_name)
    text = docx2txt.process(full_file_name)

    return fullText


def read_pdf(full_file_name):

    pdfFileObj = open(full_file_name, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    fullText = ""
    for page in range(pdfReader.numPages):
        fullText = fullText + pdfReader.getPage(page).extractText()
    return fullText

def load_data(path):
    d = []
    max_nb_words = 0
    for root, dirs, files in os.walk(path):
        if(dirs == []):
            #print(root)
            label = root.split("\\")[-1]
            print(label)
            for file in files:

               #f = open(root + "\\" + file, 'rb')
                print(file)
                full_file_name = root + "\\" + file

                if(os.path.splitext(file)[1] == ".docx"):
                    text = read_docx(full_file_name)
                elif(os.path.splitext(file)[1] == ".doc"):
                    full_file_name_doc = convert_doc2docx(full_file_name)
                    text = read_docx(full_file_name_doc)
                    os.remove(full_file_name_doc)

                elif(os.path.splitext(file)[1]==".pdf"):
                    text = read_pdf(full_file_name)
                else:
                    print("Unsupported file", file)

                d.append({'text': text, 'label': label , 'file': file})


    df = pd.DataFrame().from_dict(d)
    #df['label'] = LabelEncoder().fit_transform(df['label'])
    y = LabelBinarizer().fit_transform(LabelEncoder().fit_transform(df['label']))
    df['label'] = list(y)
    # fit LabelEncoder to the labels


    return df


path = "C:\\Users\\aelsalla\\Documents\\Valeo Documents\\Official & Mgmt\\Screening CVS\\DAS\\Screening"
df = load_data(path=path)
print(df.shape)
print(df.head(3))

from sklearn.model_selection import train_test_split
train, test = train_test_split(df, random_state=666, train_size=0.9)
#lb = LabelBinarizer().fit(df['label'])
#train_targets = lb.transform(train.label)
#test_targets = lb.transform(test.label)
train_targets = np.array(list(train['label']))
test_targets = np.array(list(test['label']))
#PROCESS TEXT: RAW
print("Text to seq process...")
print("   Fitting tokenizer...")
from keras.preprocessing.text import Tokenizer
raw_text = df.text.str.lower()

tok_raw = Tokenizer()
tok_raw.fit_on_texts(raw_text)
print("   Transforming text to seq...")
train["seq_text"] = tok_raw.texts_to_sequences(train.text.str.lower())
test["seq_text"] = tok_raw.texts_to_sequences(test.text.str.lower())


print('[{}] Finished PROCESSING TEXT DATA...'.format(time.time() - start_time))

#EMBEDDINGS MAX VALUE
#print(np.max(train.seq_text.max()))
#print(np.max(test.seq_text.max()))

print('[{}] Finished EMBEDDINGS MAX VALUE...'.format(time.time() - start_time))


#KERAS DATA DEFINITION
from keras.preprocessing.sequence import pad_sequences
#MAX_NB_WORDS_PER_DOC = 10000
MAX_NB_WORDS_PER_DOC = max([max(train["seq_text"].apply(lambda x: len(x))), max(train["seq_text"].apply(lambda x: len(x)))])
print(MAX_NB_WORDS_PER_DOC)
def get_keras_data(dataset):
    X = {
        'text': pad_sequences(dataset.seq_text, maxlen=MAX_NB_WORDS_PER_DOC)
    }
    return X

X_train = get_keras_data(train)
#X_valid = get_keras_data(dvalid)
X_test = get_keras_data(test)

print(X_train['text'].max())
print(X_test['text'].max())
#MAX_TEXT = max(X_train['text'].max, X_test['text'].max)
MAX_TEXT = max(tok_raw.word_index.values()) + 2
print(MAX_TEXT)
print('[{}] Finished DATA PREPARARTION...'.format(time.time() - start_time))



#KERAS MODEL DEFINITION
from keras.layers import Input, Dropout, Dense, BatchNormalization, \
    Activation, concatenate, GRU, Embedding, Flatten
from keras.models import Model
from keras.callbacks import ModelCheckpoint, Callback, EarlyStopping#, TensorBoard
from keras import backend as K
from keras import optimizers
from keras import initializers


dr = 0.25
nb_classes = 3
def get_model():
    #params
    dr_r = dr
    
    #Inputs
    text = Input(shape=[X_train["text"].shape[1]], name="text")

    #Embeddings layers
    emb_size = 60
    

    #emb_text = Embedding(MAX_TEXT, emb_size)(text)
    emb_text = Embedding(MAX_TEXT, emb_size)(text)

    rnn_layer1 = GRU(16) (emb_text)

    #main layer
    #main_l = concatenate([rnn_layer1])
    main_l = rnn_layer1
    main_l = Dropout(0.25)(Dense(512,activation='elu') (main_l))
    main_l = Dropout(0.2)(Dense(64,activation='elu') (main_l))
    main_l = Dropout(0.2)(Dense(64, activation='elu')(main_l))
    #main_l = Dropout(0.2)(Dense(nb_classes, activation='elu')(main_l))
    #output
    output = Dense(nb_classes,activation="softmax") (main_l)
    
    #model
    model = Model(text, output)
    #optimizer = optimizers.RMSprop()
    optimizer = optimizers.Adam()
    model.compile(loss='categorical_crossentropy',
                  optimizer=optimizer,
                  metrics=['accuracy'])
    return model

exp_decay = lambda init, fin, steps: (init/fin)**(1/(steps-1)) - 1

print('[{}] Finished DEFINING MODEL...'.format(time.time() - start_time))


gc.collect()

#FITTING THE MODEL
epochs = 2
BATCH_SIZE = 512 * 3
steps = int(len(X_train['text'])/BATCH_SIZE) * epochs
lr_init, lr_fin = 0.009, 0.006
lr_decay = exp_decay(lr_init, lr_fin, steps)
log_subdir = '_'.join(['ep', str(epochs),
                    'bs', str(BATCH_SIZE),
                    'lrI', str(lr_init),
                    'lrF', str(lr_fin),
                    'dr', str(dr)])

model = get_model()
print(model.summary())
K.set_value(model.optimizer.lr, lr_init)
K.set_value(model.optimizer.decay, lr_decay)

history = model.fit(X_train, train_targets
                    , epochs=epochs
                    , batch_size=BATCH_SIZE
                    , validation_split=0.01
                    #, callbacks=[TensorBoard('./logs/'+log_subdir)]
                    , verbose=1
                    )
print('[{}] Finished FITTING MODEL...'.format(time.time() - start_time))
#EVLUEATE THE MODEL ON DEV TEST
loss, acc = model.evaluate(X_test, test_targets, batch_size=BATCH_SIZE)

print('[{}] Finished predicting test set...'.format(time.time() - start_time))
print("Test loss:", loss, ', test acc:', acc)
