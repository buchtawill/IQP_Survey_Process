# coding=utf-8
import matplotlib                       # type: ignore
from matplotlib import image            # type: ignore
from matplotlib import pyplot as plt    # type: ignore
import pandas as pd                     # type: ignore
import numpy as np                      # type: ignore
import os
import shutil
import math

'''

Input excel file: 2 sheets, first sheet english responses, second sheet chinese
Data Columns
------------

    Timestamp

    ___________ Conditionals ___________
    Country         ['United States', 'Taiwan', 'Other'] X [臺灣, 美國, other]
    Local Resident  ['Yes' / 'No'] X [對, 不]
    Age             ['18-29', '30-44', '45-59', '60+'] [both versions]
    Gender          ['Male', 'Female', 'Do not wish to say', other] X [男性, 女性, 不想說]
    View mode       ['Smartphone', 'Tablet', 'Desktop / Laptop'] X [手機, 平板, 桌上型電腦/筆記型電腦]

    ___________ Questions ___________
    Overall Experience.......................['Very Poor', 'Poor', 'Neutral', 'Good',  'Very Good']
    Ease of Navigation.......................['Very Difficult', 'Difficult', 'Neutral', 'Easy', 'Very Easy']
    Enjoyed presentation of material.........['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
    Load Speed Response......................['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied']  

    How likely to recommend to friend........['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely']
    This website captures culture of Shilin..['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
    How likely are you to use when visiting..['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely']
    Showcases impacts of modernization.......['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
    Impact of website on survey taker........['None', ---, ----, ---, 'Very Impactful']
    
    Favorite Story      ['Wu, Jian-Hong', 'Wang, Chun-Kai', 'Lily']
    Why?                __________  Short Answer  ______

    Favorite Place      ['Zhishanyan Huiji Temple', 'Taipei MRT', 'Shilin Elementary School', 'Shilin Paper Mill', 'Shilin Architecture']
    Why?                __________  Short Answer  ______

    Something you liked about the website
    Something you disliked about the website
    See anything added?
    Any other comments?

    Make graphs for all questions across all groups
    Make bar chart of ages of respondents
    Make bar chart of Gender of respondents
    Make graphs of averages across the age groups
    Show bar charts of male and female responses
'''
'''
country_options =    ['United States', 'Taiwan', 'Other']
resident_options =   ['Yes', 'No']
age_options =        ['18-29', '30-44', '45-59', '60+']
gender_options =     ['Male', 'Female', 'Do not wish to say', 'other']
view_options =       ['Smartphone', 'Tablet', 'Desktop / Laptop']

experience_options = ['Very Poor', 'Poor', 'Neutral', 'Good',  'Very Good']
navigation_options = ['Very Difficult', 'Difficult', 'Neutral', 'Easy', 'Very Easy']
present_options =    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
speed_options =      ['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied']  

friend_options =     ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely']
culture_options =    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
visiting_options =   ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely']
modernize_options =  ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
impact_options   =   ['None', '---', '---', '---', 'Very Impactful']

fav_story_options =  ['Wu, Jian-Hong (吳儉鴻)', 'Wang, Chun-Kai (王俊凱)', 'Lily (莉莉)']
fav_story_reason = None
fav_place_options =  ['Zhishanyan Huiji Temple (芝山巖惠濟宮)', 'Taipei MRT (台北捷運)', 'Shilin Elementary School (士林國小)', 'Shilin Paper Mill (士林紙廠)', 'Shilin Architecture (士林建築)']

'''


response_options = {
'country' :                 ['United States', 'Taiwan', 'Other'],
'local resident' :          ['Yes', 'No'],
'age' :                     ['18-29', '30-44', '45-59', '60+'],
'gender' :                  ['Male', 'Female', 'Do not wish to say', 'other'],
'view mode' :               ['Smartphone', 'Tablet', 'Desktop / Laptop'],

'overall experience' :      ['Very Poor', 'Poor', 'Neutral', 'Good',  'Very Good'],
'ease of navigation' :      ['Very Difficult', 'Difficult', 'Neutral', 'Easy', 'Very Easy'],
'enjoyed presentation' :    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
'load speed' :              ['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied'],

'recommend to friend' :     ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
'captures culture' :        ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
'use when visiting' :       ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
'showcases modernization' : ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
'impactfulness'   :         ['None', 'Little', 'Neutral', 'Impactful', 'Very Impactful'],

'favorite story' :          ['Wu, Jian-Hong (吳儉鴻)', 'Wang, Chun-Kai (王俊凱)', 'Lily (莉莉)'],
'favorite place' :          ['Zhishanyan Huiji Temple (芝山巖惠濟宮)', 'Taipei MRT (台北捷運)', 
                             'Shilin Elementary School (士林國小)', 'Shilin Paper Mill (士林紙廠)', 
                             'Shilin Architecture (士林建築)'],
}


def cndf_to_en(cn: pd.DataFrame) -> pd.DataFrame:
    #country you live in
    for i in range(len(cn['country'])):
        ans = cn['country'][i]
        if(ans == '臺灣'):
            cn.loc[i, 'country'] = 'Taiwan'
        
        elif(ans =='美國'):
            cn.loc[i, 'country'] = 'United States'
        #else dont care, keep original (other)

    #is a local resident
    for i in range(len(cn['local resident'])):
        #print(cn.loc[i, 'local resident'])
        ans = cn.loc[i, 'local resident']
        if(ans =='對'):
            cn.loc[i, 'local resident'] = 'Yes'
        else:
            cn.loc[i, 'local resident'] = 'No'

    #don't need to do age row

    #gender
    for i in range(len(cn['gender'])):
        ans = cn['gender'][i]
        if(ans == '男性'):
            cn.loc[i, 'gender'] = 'Male'
        
        elif(ans =='女性'):
            cn.loc[i, 'gender'] = 'Female'

        elif(ans =='不想說'):
            cn.loc[i, 'gender'] = 'Do not wish to say'
        #else dont care, keep original (other)

    #view mode
    for i in range(len(cn['view mode'])):
        ans = cn['view mode'][i]
        if(ans == '手機'):
            cn.loc[i, 'view mode'] = 'Smartphone'
        
        elif(ans =='平板'):
            cn.loc[i, 'view mode'] = 'Tablet'

        elif(ans =='桌上型電腦/筆記型電腦'):
            cn.loc[i, 'view mode'] = 'Desktop / Laptop'
    
    return cn

IMAGE_TRANSPARENCY = False
IMAGE_TEXT_COLOR = 'black'
IMAGE_AXIS_COLOR = 'black'
OUTPUT_IMAGE_LOCATION = "./out/"
SAVE = True
DPI = 300

def setMatplotParams():
    matplotlib.rcParams['text.color'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['axes.labelcolor'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['xtick.color'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['ytick.color'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['axes.edgecolor'] = IMAGE_AXIS_COLOR

#sort x and y, according to values in y
def selectionSort(x, y, reverse = False):
    #selection sort bc it's OP for small sized arrays
    
    if(reverse is False):
        for i in range(len(y)):
            smallest_i = i
            for j in range(i+1, len(y)):
                if(y[j] < y[smallest_i]):
                    smallest_i = j
            
            (x[i], x[smallest_i]) = (x[smallest_i], x[i])
            (y[i], y[smallest_i]) = (y[smallest_i], y[i])
            
    else:
        for i in range(len(y)):
            biggest_i = i
            for j in range(i+1, len(y)):
                if(y[j] > y[biggest_i]):
                    biggest_i = j
            
            (x[i], x[biggest_i]) = (x[biggest_i], x[i])
            (y[i], y[biggest_i]) = (y[biggest_i], y[i])

def plot_other_bar(df: pd.DataFrame, title: str, column:str, show=False, sort='up'):
    plt.cla()
    
    d = dict()
    x = []
    y = []

    for option in response_options[column]:
        print(option)
        d[option] = int(0)
    
    for i in range(len(df[column])):
        row = df.loc[i, column]
        print(row)


    for row in df[column]:
        #if a new response, add it to the dict (country or gender)
        if(row not in d):
            print(f"its not there bruv: {row}")
            d[row] = int(1)
        else:    
            d[row] = d[row] + int(1)
    

    #every i in d.items() is a tuple of (key, val)]
    #print(d.items())
    for i in d.items():
        x.append(i[0])
        y.append(i[1])
    
    if(column == 'favorite story'):
        x = ['Wu, Jian-Hong', 'Wang, Chun-Kai', 'Lily']

    elif(column == 'favorite place'):
        x = ['Huiji Temple', 'Taipei MRT', 'Shilin Elementary', 'Shilin Paper Mill', 'Shilin Architecture']
    
    if(sort == 'up'):
        selectionSort(x, y, False)
    elif(sort =='down'):
        selectionSort(x, y, True)
        
    
    plt.bar(x, y, align='center')
    plt.xticks(rotation=35, ha='left')
    plt.title(f"{title} (n={len(df[column])})")
    plt.ylabel("Count")
    plt.xlabel("")

    # Add labels to the bars
    for i, v in enumerate(y):
        plt.text(i, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning

    if(SAVE):
        plt.savefig(OUTPUT_IMAGE_LOCATION + column + '.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    if(show):
        plt.show()


def plot_likert(df: pd.DataFrame, title: str, column:str, show=False):
    plt.cla()
    
    x = response_options[column] # 'unlikely', 'neutral', ....
    y = np.zeros(len(x), dtype=int)

    for entry in df[column]:
        if(not math.isnan(entry)):
            entry = int(entry)
            y[entry-1] += 1
    
    total_resp = np.sum(y)
    sum = 0
    for i in range(len(y)):
        sum += y[i] * (i+1)

    avg = float(sum) / total_resp
    ax = plt.bar(x, y)
    
    plt.xticks(rotation=35, ha='center')
    plt.title(f"{title} (n={len(df[column])}, avg={round(avg,1)})")
    plt.ylabel("Count")
    plt.xlabel("")

    # Add labels to the bars
    for i, v in enumerate(y):
        plt.text(i, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning
    
    if(SAVE):
        plt.savefig(OUTPUT_IMAGE_LOCATION + column + '.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    
    if(show):
        plt.show()

if __name__ == '__main__':

    source = 'responses.xlsx'
    column_labels = ['timestamp', 'country', 'local resident', 'age', 'gender', 'view mode',                                #options
               'overall experience', 'ease of navigation', 'enjoyed presentation', 'load speed',                            #likert
               'recommend to friend', 'captures culture', 'use when visiting', 'showcases modernization', 'impactfulness',  #likert
               'favorite story', 'why story', 'favorite place', 'why place',                                                #option, short ans
               'something you liked', 'something you disliked', 'want to add anything', 'other comments']                   #long answer
    
    xls = pd.ExcelFile(source)

    cn = pd.read_excel(source, 'Chinese', header=None, names=column_labels)
    en = pd.read_excel(source, 'English', header=None, names=column_labels)

    cn = cndf_to_en(cn)
    data = pd.concat([en, cn], ignore_index=True)

    i = data[((data.country != 'United States'))].index
    #data = data.drop(i)

    if(SAVE):
        try:
            shutil.rmtree(OUTPUT_IMAGE_LOCATION)
        except FileNotFoundError:
            pass

        if not os.path.exists(OUTPUT_IMAGE_LOCATION):
            os.makedirs(OUTPUT_IMAGE_LOCATION)
        
    setMatplotParams()
    
    #              df    plot title                     column title                    sort direction (default = up)
    plot_other_bar(data, "Favorite Story",              'favorite story', show=True,    sort='down')
    plot_other_bar(data, "Favorite \'Place\' Webpage",  'favorite place', show=True,    sort='down')

    '''plot_other_bar(data, "Viewing Device",          'view mode',        sort='up')
    plot_other_bar(data, 'Are You a Local?',        'local resident',   sort='no')
    plot_other_bar(data, "Age",                     'age',              sort='no')
    plot_other_bar(data, "Respondents' Country",    'country',          sort='up')
    plot_other_bar(data, "Respondents' Gender",     'gender',           sort='down')

    
    
    
    #likert-not sorted
    plot_likert(data, "Overall Experience",                             'overall experience')
    plot_likert(data, "Easy to Navigate?",                              'ease of navigation')
    plot_likert(data, "Enjoyed Presentation of Material",               'enjoyed presentation')
    plot_likert(data, "Reaction to Load Speed",                         'load speed')
    plot_likert(data, "Would Recommend to a Friend",                    'recommend to friend')
    plot_likert(data, "\'This Website Captures Shilin\'s Culture\'",    'captures culture')
    plot_likert(data, "Would Use Website When Visiting Shilin",         'use when visiting')
    plot_likert(data, "\'This Website Captures Modernization\'",        'showcases modernization')
    plot_likert(data, "How Much of an Impact Do You Feel?",             'impactfulness')'''
    
    
    
    
    
