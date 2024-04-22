# coding=utf-8
import matplotlib                       # type: ignore
from matplotlib import image            # type: ignore
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
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
                        ['Wu, Jian-Hong (吳儉鴻)', 'Wang, Chun-Kai (王俊凱)', 'Lily (莉莉)']
    Why?                __________  Short Answer  ______

    Favorite Place      ['Zhishanyan Huiji Temple', 'Taipei MRT', 'Shilin Elementary School', 'Shilin Paper Mill', 'Shilin Architecture']
                        ['Zhishanyan Huiji Temple (芝山巖惠濟宮)', 'Taipei MRT (台北捷運)', 'Shilin Elementary School (士林國小)', 'Shilin Paper Mill (士林紙廠)', 'Shilin Architecture (士林建築)']
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
    
    Colors: https://www.color-hex.com/color-palette/96986
'''

response_options = {
    'country' :                 ['United States', 'Taiwan'], #or other
    'local resident' :          ['Yes', 'No'],
    'age' :                     ['18-29', '30-44', '45-59', '60+'],
    'gender' :                  ['Male', 'Female', 'Do not wish to say'], #or other
    'view mode' :               ['Smartphone', 'Tablet', 'Desktop / Laptop'],

    'overall experience' :      ['Very Poor', 'Poor', 'Neutral', 'Good',  'Very Good'],
    'ease of navigation' :      ['Very Difficult', 'Difficult', 'Neutral', 'Easy', 'Very Easy'],
    'enjoyed presentation' :    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
    'load speed' :              ['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied'],

    'recommend to friend' :     ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
    'captures culture' :        ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
    'use when visiting' :       ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
    'showcases modernization' : ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
    'impactfulness'   :         ['None', 'Very Little', 'Somewhat', 'Impactful', 'Very Impactful'],

    'favorite story' :          ['Wu, Jian-Hong', 'Wang, Chun-Kai', 'Lily'],
    'favorite place':           ['Huiji Temple', 'Taipei MRT', 'Shilin Elementary', 'Shilin Paper Mill', 'Shilin Architecture']
}

column_labels = ['timestamp', 'country', 'local resident', 'age', 'gender', 'view mode',                                    #options
               'overall experience', 'ease of navigation', 'enjoyed presentation', 'load speed',                            #likert
               'recommend to friend', 'captures culture', 'use when visiting', 'showcases modernization', 'impactfulness',  #likert
               'favorite story', 'why story', 'favorite place', 'why place',                                                #option, short ans
               'something you liked', 'something you disliked', 'want to add anything', 'other comments']                   #long answer

#convert chinese responses to english
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

#given a pandas column, return x and y with x being labels, y being number of responses
def columnToXY(df: pd.DataFrame, column: str):
    d = dict()
    x = []
    y = []

    for option in response_options[column]:
        d[option] = int(0)
    
    #count up all the occurences of each response option
    for row in df[column]:
        #if a new response, add it to the dict (country or gender 'other')
        if(row not in d):
            print(f"Adding new \'other\' option for {column}: \'{row}\'")
            d[row] = int(1)
        else:    
            d[row] = d[row] + int(1)
    
    #Put the dict into [x,y] to sort later and graph easier
    for i in d.items():
        x.append(i[0])
        y.append(int(i[1]))
        
    return x,y 

def plot_pie_other(df: pd.DataFrame, title: str, column:str, show=False):
    plt.cla()
    x, y = columnToXY(df, column)
    newx = []
    newy = []
    
    for i in range(len(y)):
        if(y[i] > 0):
            newx.append(x[i])
            newy.append(y[i])
    x = newx
    y = newy
    #selectionSort(x, y)
    
    d = pd.DataFrame([x,y])
    d.to_excel(f"{OUTPUT_IMAGE_LOCATION}{title}.xlsx")
      
    plt.pie(y, labels=x, autopct='%1.1f%%', radius=1.25, 
            #colors=['#4deeea', '#74ee15','#ffaa00', '#ffe700', '#f000ff', '#00aaff'])
            #colors=['#00aaff', '#aa00ff', '#ff00aa', '#ffaa00', '#aaff00'])
            colors=['#ffaa00', '#aaff00', '#ff00aa', '#aa00ff', '#00aaff'])
    
    plt.title(f"{title}, group={OUTPUT_IMAGE_LOCATION[2:-1]} (n={np.sum(y)})", y=1.08)
    
    if(SAVE):
        plt.savefig(OUTPUT_IMAGE_LOCATION + column + '.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    if(show):
        plt.show()


def plot_other_bar(df: pd.DataFrame, title: str, column:str, show=False, sort='up'):
    plt.cla()
    
    x, y = columnToXY(df, column)
    
    if(sort == 'up'):
        selectionSort(x, y, False)
    elif(sort =='down'):
        selectionSort(x, y, True)
        
    
    plt.bar(x, y, align='center')
    plt.xticks(rotation=15, ha='center')
    plt.title(f"{title}, group={OUTPUT_IMAGE_LOCATION[2:-1]} (n={len(df[column])})")
    plt.ylabel("Count")
    plt.gca().yaxis.set_major_formatter(ticker.FormatStrFormatter('%d'))
    plt.xlabel("")

    # Add labels to the bars
    for i, v in enumerate(y):
        plt.text(i, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning

    if(SAVE):
        plt.savefig(OUTPUT_IMAGE_LOCATION + column + '.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    if(show):
        plt.show()

def otherTwoGroups(df1: pd.DataFrame, label1:str, df2: pd.DataFrame, label2: str, plottitle: str, column:str, show=False):
    
    plt.cla()
    barSpacing = 0.2
    x, y1 = columnToXY(df1, column)
    x, y2 = columnToXY(df2, column)
    
    x_axis = np.arange(len(x)) 
    plt.bar(x_axis - barSpacing, y1, 0.4, label = label1) 
    plt.bar(x_axis + barSpacing, y2, 0.4, label = label2) 
    
    plt.xticks(x_axis, x, rotation=15, ha='center')
    #plt.title(f"{plottitle}, group={OUTPUT_IMAGE_LOCATION[2:-1]} (n={len(df[column])}, avg={round(avg,1)})")
    plt.title(f"{plottitle} (n={len(df1[column])+len(df2[column])})")
    plt.ylabel("Count")
    plt.xlabel("")
    plt.legend() 
    plt.gca().yaxis.set_major_formatter(ticker.FormatStrFormatter('%d'))

    # Add labels to the bars
    for i, v in enumerate(y1):
        plt.text(i-barSpacing, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning
    for i, v in enumerate(y2):
        plt.text(i+barSpacing, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning
    
    if(SAVE):
        plt.savefig(f'{OUTPUT_IMAGE_LOCATION}{plottitle}.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    
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
    
    plt.xticks(rotation=15, ha='center')
    plt.title(f"{title}, group={OUTPUT_IMAGE_LOCATION[2:-1]} (n={len(df[column])}, avg={round(avg,1)})")
    plt.ylabel("Count")
    plt.xlabel("")

    # Add labels to the bars
    for i, v in enumerate(y):
        plt.text(i, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning
    
    if(SAVE):
        plt.savefig(OUTPUT_IMAGE_LOCATION + column + '.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    
    if(show):
        plt.show()


def likertTwoGroups(df1: pd.DataFrame, label1:str, df2: pd.DataFrame, label2: str, plottitle: str, column:str, show=False):
    plt.cla()
    
    barSpacing = 0.2
    
    x = response_options[column] # 'unlikely', 'neutral', ....
    y1 = np.zeros(len(x), dtype=int)
    y2 = np.zeros(len(x), dtype=int)
    
    for entry in df1[column]:
        #print(entry)
        if(not math.isnan(entry)):
            entry = int(entry)
            y1[entry-1] += 1
    #print('----------------')
    for entry in df2[column]:
        #print(entry)
        if(not math.isnan(entry)):
            entry = int(entry)
            y2[entry-1] += 1
    
    x_axis = np.arange(len(x)) 
    plt.bar(x_axis - barSpacing, y1, 0.4, label = label1) 
    plt.bar(x_axis + barSpacing, y2, 0.4, label = label2) 
    
    plt.xticks(x_axis, x, rotation=15, ha='center')
    #plt.title(f"{plottitle}, group={OUTPUT_IMAGE_LOCATION[2:-1]} (n={len(df[column])}, avg={round(avg,1)})")
    plt.title(f"{plottitle} (n={len(df1[column])+len(df2[column])})")
    plt.ylabel("Count")
    plt.xlabel("")
    plt.legend() 
    plt.gca().yaxis.set_major_formatter(ticker.FormatStrFormatter('%d'))

    # Add labels to the bars
    for i, v in enumerate(y1):
        plt.text(i-barSpacing, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning
    for i, v in enumerate(y2):
        plt.text(i+barSpacing, v + 0.1, str(v), ha='center', va='bottom')  # Adjust the + 0.1 and 'bottom' as needed for positioning
    
    if(SAVE):
        plt.savefig(f'{OUTPUT_IMAGE_LOCATION}{plottitle}.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight', dpi=DPI)
    
    if(show):
        plt.show()

def fixFavoriteStory(data: pd.DataFrame) -> pd.DataFrame:
    #fix favorite story column (parenthesis bug)
    for i in range(len(data['favorite story'])):
        row = data.loc[i, 'favorite story']
        #print(row)
        if('Lily' in row):
            data.loc[i, 'favorite story'] = 'Lily'
        elif ('Wang' in row):
            data.loc[i, 'favorite story'] = 'Wang, Chun-Kai'
        elif ('Jian-Hong' in row):
            data.loc[i, 'favorite story'] = 'Wu, Jian-Hong'
    
    return data

def setOtherCountries(data: pd.DataFrame) -> pd.DataFrame:
    for i in range(len(data['country'])):
        row = data.loc[i, 'country']
        if(row != 'United States' and row != 'Taiwan'):
            data.loc[i, 'country'] = 'Other'
    return data

def fixFavoritePlace(data: pd.DataFrame) -> pd.DataFrame:
    #fix favorite place column (remove chinese characters)
    for i in range(len(data['favorite place'])):
        row = data.loc[i, 'favorite place']
        #print(row)
        if('Huiji' in row):
            data.loc[i, 'favorite place'] = 'Huiji Temple'
        elif ('MRT' in row):
            data.loc[i, 'favorite place'] = 'Taipei MRT'
        elif ('Elementary' in row):
            data.loc[i, 'favorite place'] = 'Shilin Elementary'
        elif ('Paper' in row):
            data.loc[i, 'favorite place'] = 'Shilin Paper Mill'
        elif ('Arch' in row):
            data.loc[i, 'favorite place'] = 'Shilin Architecture'
    
    return data

def default_graphs(data: pd.DataFrame):
    #              df    plot title                     column title        sort direction (default = up)
    plot_other_bar(data, "Favorite Story",              'favorite story',   sort='no')
    plot_other_bar(data, "Favorite \'Place\' Webpage",  'favorite place',   sort='down')

    plot_other_bar(data, "Viewing Device",              'view mode',        sort='up')
    plot_other_bar(data, 'Are You a Local?',            'local resident',   sort='no')
    plot_other_bar(data, "Age",                         'age',              sort='no')
    plot_other_bar(data, "Respondents' Country",        'country',          sort='up')
    plot_other_bar(data, "Respondents' Gender",         'gender',           sort='down')

    #likert - not sorted
    plot_likert(data, "Overall Experience",                             'overall experience')
    plot_likert(data, "Easy to Navigate?",                              'ease of navigation')
    plot_likert(data, "Enjoyed Presentation of Material",               'enjoyed presentation')
    plot_likert(data, "Reaction to Load Speed",                         'load speed')
    plot_likert(data, "Would Recommend to a Friend",                    'recommend to friend')
    plot_likert(data, "\'This Website Captures Shilin\'s Culture\'",    'captures culture')
    plot_likert(data, "Would Use Website When Visiting Shilin",         'use when visiting')
    plot_likert(data, "\'This Website Captures Modernization\'",        'showcases modernization')
    plot_likert(data, "How Much of an Impact Do You Feel?",             'impactfulness')
    
    plot_pie_other(data, "Country", 'country')
    plot_pie_other(data, "Favorite_Story", 'favorite story')
    plot_pie_other(data, "Favorite_Place", 'favorite place')
    plot_pie_other(data, "View_Mode", 'view mode')
    plot_pie_other(data, "Are_you_local", 'local resident')
    plot_pie_other(data, "Age", 'age')
    plot_pie_other(data, "Gender", 'gender')

def two_group_likerts(df1: pd.DataFrame, t1: str, df2: pd.DataFrame, t2: str):
    
    likertTwoGroups(df1, t1, df2, t2, f"Overall Experience ({t1} and {t2})", 'overall experience')
    likertTwoGroups(df1, t1, df2, t2, f"Ease of Navigation ({t1} and {t2})", 'ease of navigation')
    likertTwoGroups(df1, t1, df2, t2, f"Enjoyed Presentation ({t1} and {t2})", 'enjoyed presentation')
    likertTwoGroups(df1, t1, df2, t2, f"Response to Load Speed ({t1} and {t2})", 'load speed')
    
    likertTwoGroups(df1, t1, df2, t2, f"Recommend to a Friend ({t1} and {t2})", 'recommend to friend')
    likertTwoGroups(df1, t1, df2, t2, f"Captures Culture ({t1} and {t2})", 'captures culture')
    likertTwoGroups(df1, t1, df2, t2, f"Showcases Modernization ({t1} and {t2})", 'showcases modernization')
    likertTwoGroups(df1, t1, df2, t2, f"Impactfulness ({t1} and {t2})", 'impactfulness')
    likertTwoGroups(df1, t1, df2, t2, f"Would Use When Visiting ({t1} and {t2})", 'use when visiting')
    

def taiwaneseOnly(data):
    i = data[((data.country != 'Taiwan'))].index
    data = data.drop(i)
    data.reset_index()
    return data

def otherCountriesOnly(data):
    i = data[((data.country == 'Taiwan'))].index
    data = data.drop(i)
    data.reset_index()
    return data

def youngGroup(data):
    i = data[((data.age != '18-29'))].index
    data = data.drop(i)
    data.reset_index()
    return data

def oldGroups(data):
    i = data[((data.age == '18-29'))].index
    data = data.drop(i)
    data.reset_index()
    return data

def femaleOnly(data):
    i = data[((data.gender != 'Female'))].index
    data = data.drop(i)
    data.reset_index()
    return data

def maleOnly(data):
    i = data[((data.gender != 'Male'))].index
    data = data.drop(i)
    data.reset_index()
    return data

IMAGE_TRANSPARENCY = False
IMAGE_TEXT_COLOR = 'black'
IMAGE_AXIS_COLOR = 'black'
OUTPUT_IMAGE_LOCATION = "./visitors_vs_locals/"
SAVE = True
DPI = 300

if __name__ == '__main__':

    source = 'responses.xlsx'
    xls = pd.ExcelFile(source)

    cn = pd.read_excel(source, 'Chinese', header=None, names=column_labels)
    en = pd.read_excel(source, 'English', header=None, names=column_labels)
    cn = cndf_to_en(cn)
    data = pd.concat([en, cn], ignore_index=True)
    data = fixFavoritePlace(data)   #remove chinese characters and fix parenthesis bug
    data = fixFavoriteStory(data)   #remove chinese characters and fix parenthesis bug
    data = setOtherCountries(data)  #countries other than US and Taiwan will be 'other'
    
    print(f"Total responses: {len(data['country'])}")
    
    guys  = maleOnly(data)
    girls = femaleOnly(data)
    
    taiwanese = taiwaneseOnly(data)
    visitors  = otherCountriesOnly(data)
    
    old = oldGroups(data)
    young = youngGroup(data)
    
    if(SAVE):
        try:
            shutil.rmtree(OUTPUT_IMAGE_LOCATION)
        except FileNotFoundError:
            pass
        if not os.path.exists(OUTPUT_IMAGE_LOCATION):
            os.makedirs(OUTPUT_IMAGE_LOCATION)
        
    setMatplotParams()
    
    #likertTwoGroups(guys, 'Male', girls, 'Female', "Would Recommend to a Friend (Female and Male)", 'recommend to friend')
    #otherTwoGroups(guys, 'Male', girls, "Female", "Favorite Story (Female and Male)", 'favorite story')
    #otherTwoGroups(guys, 'Male', girls, "Female", "Favorite place (Female and Male)", 'favorite place')
    
    #default_graphs(visitors)
    two_group_likerts(taiwanese, 'Taiwanese', visitors, 'Not Taiwanese')
    
