# Packages and imports
import nltk

nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import string
import os
import PIL


# Import file
os.chdir('C:/Users/jiasi/Desktop/Course Materials/AA500/TIPS Project') # Change it to your own working directory before running
dic = pd.read_excel(r'TIPS.xlsx', sheet_name=None)


# Make a list of all comments and remove NA's in excel
years = list(dic.keys())

comm_raw = []
for i in years:
    df = dic[i]
    vallist = df.values.tolist()
    for j in vallist:
        comm_raw.extend(j)

comm_val = [x for x in comm_raw if pd.isnull(x) == False]

print("We have the following years' comments")
print(years)


# Separate each year's comments
comm_raw_2021 = [] ## Year 2021
vallist = dic['2021'].values.tolist()
for j in vallist:
    comm_raw_2021.extend(j)
comm_val_2021 = [x for x in comm_raw_2021 if pd.isnull(x) == False]

comm_raw_2020 = [] ## Year 2020
vallist = dic['2020'].values.tolist()
for j in vallist:
    comm_raw_2020.extend(j)
comm_val_2020 = [x for x in comm_raw_2020 if pd.isnull(x) == False]
    
comm_raw_2019 = [] ## Year 2019
vallist = dic['2019'].values.tolist()
for j in vallist:
    comm_raw_2019.extend(j)
comm_val_2019 = [x for x in comm_raw_2019 if pd.isnull(x) == False]    

comm_raw_2018 = [] ## Year 2018
vallist = dic['2018'].values.tolist()
for j in vallist:
    comm_raw_2018.extend(j)
comm_val_2018 = [x for x in comm_raw_2018 if pd.isnull(x) == False]    

comm_raw_2017 = [] ## Year 2017
vallist = dic['2017'].values.tolist()
for j in vallist:
    comm_raw_2017.extend(j)
comm_val_2017 = [x for x in comm_raw_2017 if pd.isnull(x) == False]    

comm_raw_2016 = [] ## Year 2016
vallist = dic['2016'].values.tolist()
for j in vallist:
    comm_raw_2016.extend(j)
comm_val_2016 = [x for x in comm_raw_2016 if pd.isnull(x) == False]


# Headcount
print("There are", len(comm_val), "valid comments in total.")
for i in years:
    print("In year", i, "there are" ,len(globals()['comm_val_%s' % i]) ,"valid comments.")
    


# Make a list of stop words which are not meaningful (like "WHICH")
stop_words = nltk.corpus.stopwords.words('english')

from nltk.stem import PorterStemmer # stemming function
from nltk.stem.wordnet import WordNetLemmatizer # lemmatization functions
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator # wordcloud functions

## Word cloud function defined
def make_cloud(k, m, s, f): # where k is the input validated comments set
                            # m is the png file name
                            # s is the maximum font size
                            # f is the name of the figure exported
    
    global stop_words
    
    # Break comments into tokens to create a nested list
    tkn = [] 
    for i in k:
        x = nltk.word_tokenize(i)
        tkn.append(x)
    
    tkn_fil = []
    for i in tkn:
        temp_list = []
        for j in i:
            if j not in stop_words:
                temp_list.append(j)
        tkn_fil.append(temp_list)
    
    # Word stemming
    stem = PorterStemmer()

    tkn_stem = []
    for i in tkn_fil:
        temp_list = []
        for j in i:
            temp_list.append(stem.stem(j))
        tkn_stem.append(temp_list)
    
    # Word lemmatization
    lem = WordNetLemmatizer()

    tkn_lem = []
    for i in tkn_fil:
        temp_list = []
        for j in i:
            temp_list.append(lem.lemmatize(j, "v"))
        tkn_lem.append(temp_list)
    
    # Create string vectors
    str_cloud =''
    for i in tkn_lem:
        for j in i:
            str_cloud = str_cloud + ' ' + j.lower()
    
    # Make word clouds
    stop_words_cloud = list(STOPWORDS) + \
                        ["n\'t"] + \
                            list(string.ascii_lowercase) + \
                            ["non", "re", "ing", "de", "ed", \
                             "go", "make", "ll"] # Some prefix and suffix which are not our interests
    
    # Read image prototypes
    mask = np.array(PIL.Image.open(m))
    
    wordcloud = WordCloud(
                     scale = 5,
                     stopwords = stop_words_cloud,
                     background_color = 'white',
                     mode = "RGBA",
                     max_words= 500,
                     mask = mask,
                     max_font_size = s).generate(str_cloud)
    
    image_colors = ImageColorGenerator(mask)
    plt.figure(figsize = [7,7])
    plt.imshow(wordcloud.recolor(color_func=image_colors), interpolation='bilinear')
    plt.axis("off")
    plt.savefig(f, format = "png")
    plt.show()

    
# Run function and generate word clouds
path = "question.png"
filename = "2016_cloud.png"
make_cloud(comm_val_2016, path, 70, filename)

filename = "2019_cloud.png"
make_cloud(comm_val_2019, path, 70, filename)


path = "bulb.png"
filename = "2017_cloud.png"
make_cloud(comm_val_2017, path, 50, filename)

filename = "2020_cloud.png"
make_cloud(comm_val_2020, path, 50, filename)


path = "tick.png"
filename = "2018_cloud.png"
make_cloud(comm_val_2018, path, 70, filename)

filename = "2021_cloud.png"
make_cloud(comm_val_2021, path, 70, filename)


path = "NCSU.png"
filename = "total_cloud.png"
make_cloud(comm_val, path, 70, filename)
