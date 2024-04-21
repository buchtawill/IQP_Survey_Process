# coding=utf-8
import matplotlib                       # type: ignore
from matplotlib import image            # type: ignore
from matplotlib import pyplot as plt    # type: ignore
import pandas as pd                     # type: ignore
import numpy as np                      # type: ignore

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
survey = [
    ['United States', 'Taiwan', 'Other'],
    ['Yes', 'No'],
    ['18-29', '30-44', '45-59', '60+'],
    ['Male', 'Female', 'Do not wish to say', 'other'],
    ['Smartphone', 'Tablet', 'Desktop / Laptop'],

    ['Very Poor', 'Poor', 'Neutral', 'Good',  'Very Good'],
    ['Very Difficult', 'Difficult', 'Neutral', 'Easy', 'Very Easy'],
    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
    ['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied']  ,

    ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
    ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
    ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
    ['None', '---', '---', '---', 'Very Impactful'],

    ['Wu, Jian-Hong (吳儉鴻)', 'Wang, Chun-Kai (王俊凱)', 'Lily (莉莉)'],
    ['long_answer_who'],
    ['Zhishanyan Huiji Temple (芝山巖惠濟宮)', 'Taipei MRT (台北捷運)', 'Shilin Elementary School (士林國小)', 'Shilin Paper Mill (士林紙廠)', 'Shilin Architecture (士林建築)'],
    ['long_answer_place']
]

response_options = {
'country_options' :     ['United States', 'Taiwan', 'Other'],
'resident_options' :    ['Yes', 'No'],
'age_options' :         ['18-29', '30-44', '45-59', '60+'],
'gender_options' :      ['Male', 'Female', 'Do not wish to say', 'other'],
'view_options' :        ['Smartphone', 'Tablet', 'Desktop / Laptop'],

'experience_options' :  ['Very Poor', 'Poor', 'Neutral', 'Good',  'Very Good'],
'navigation_options' :  ['Very Difficult', 'Difficult', 'Neutral', 'Easy', 'Very Easy'],
'present_options' :     ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
'speed_options' :       ['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied'],

'friend_options' :      ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
'culture_options' :     ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
'visiting_options' :    ['Very Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Very Likely'],
'modernize_options' :   ['Stronly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree'],
'impact_options'   :    ['None', '---', '---', '---', 'Very Impactful'],

'fav_story_options' :   ['Wu, Jian-Hong (吳儉鴻)', 'Wang, Chun-Kai (王俊凱)', 'Lily (莉莉)'],
'fav_story_reason' :    [],
'fav_place_options' :   ['Zhishanyan Huiji Temple (芝山巖惠濟宮)', 'Taipei MRT (台北捷運)', 'Shilin Elementary School (士林國小)', 'Shilin Paper Mill (士林紙廠)', 'Shilin Architecture (士林建築)'],
'fav_place_reason' :    []
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

def setMatplotParams():
    matplotlib.rcParams['text.color'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['axes.labelcolor'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['xtick.color'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['ytick.color'] = IMAGE_TEXT_COLOR
    matplotlib.rcParams['axes.edgecolor'] = IMAGE_AXIS_COLOR

def plot_country(df: pd.DataFrame):
    bar_graph_ax = data['country'].value_counts(sort=True).plot(kind='bar')
    plt.bar_label(bar_graph_ax.containers[0])
    plt.xticks(rotation=0, ha='center')
    plt.title(f"Country of Residence (n={len(data['country'])})")
    plt.ylabel("Count")
    plt.xlabel("Country")
    #plt.savefig(OUTPUT_IMAGE_LOCATION + 'ages.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight')
    plt.show()


def plot_viewmode(df: pd.DataFrame):
    bar_graph_ax = data['view mode'].value_counts(sort=True).plot(kind='bar')
    plt.bar_label(bar_graph_ax.containers[0])
    #bar_graph_ax.set_xticklabels(['Brother', 'Bruh', 'Otto'])
    plt.xticks(rotation=0, ha='center')
    plt.title(f"View Mode (n={len(data['country'])})")
    plt.ylabel("Count")
    plt.xlabel("")
    #plt.savefig(OUTPUT_IMAGE_LOCATION + 'ages.png', transparent=IMAGE_TRANSPARENCY, bbox_inches='tight')
    plt.show()

if __name__ == '__main__':
    
    source = 'responses.xlsx'
    column_labels = ['timestamp', 'country', 'local resident', 'age', 'gender', 'view mode', 
               'overall experience', 'ease of navigation', 'enjoyed presentation', 'load speed',
               'recommend to friend', 'captures culture', 'use when visiting', 'showcases modernization', 'impactfulness',
               'favorite story', 'why story', 'favorite place', 'why place',
               'something you liked', 'something you disliked', 'want to add anything', 'other comments']
    
    xls = pd.ExcelFile(source)

    cn = pd.read_excel(source, 'Chinese', header=None, names=column_labels)
    en = pd.read_excel(source, 'English', header=None, names=column_labels)

    cn = cndf_to_en(cn)
    data = pd.concat([en, cn], ignore_index=True)

    #data.to_excel('test.xlsx')
    #i = data[((data.country == 'United States'))].index
    #data = data.drop(i)

    setMatplotParams()
    #plot_country(data)
    plot_viewmode(data)


    
    
    
