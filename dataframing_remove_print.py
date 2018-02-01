# -*- coding: utf-8 -*-

# SKKU 2017
#
# This is the script for saving the voting results of in-class activity at dataframe.
# It is optimized for the class GEDB005-41, in fall semester of 2017 in SKKU
# This dataframe is for input data of analysing communication of students in same group during class activities.

import pandas as pd
import numpy as np
import random
import matplotlib.pyplot as plt
import sys
import os
import xlrd
from openpyxl import load_workbook
import re

def ErrorCheck(df_vote):
    ## checking the mistaken of inputs
    # Does any student put the answer twice at 1st and 2nd voting?
    df_sample = df_vote.groupby(u"Student ID No. (학번)").groups
    #print(df_sample)
    for sample in df_sample:
        # get indexes of same student id
        li_index = df_sample.get(sample)
        #print(li_index)
        # Find if the number of indexes are more than 3.
        # Remove the double vote data.
        li_1st = []
        li_2nd = []
        for i in range(len(li_index)):
            # get the value of first and seconde vote.
            row = df_vote.ix[li_index[i]]
            first_vote = row["Let's vote."]
            second_vote = row["Let's vote again."]
            if not np.isnan(first_vote):
                li_1st.append({li_index[i]: first_vote})
            if not np.isnan(second_vote):
                li_2nd.append({li_index[i] : second_vote})
        #print(li_1st, li_2nd)
        # the print results are like below
        # example1. ([{13: 3.0}], [{128: 3.0}, {129: 3.0}])
        # exampel2. ([{5: 4.0}, {11: 4.0}], [{5: 4.0}, {92: 4.0}])

        # If it has double voting data and the value is same, deleted
        if len(li_1st) > 1:
            if li_1st[0].values() == li_1st[1].values():
                random_index = random.choice(li_1st).keys()[0]
                #print(random_index)
                # Check if the dataframe has the data you want to deleted.
                if random_index in df_vote.index:
                    df_vote = df_vote.drop(random_index)
                    print(random_index, "is deleted")
        if len(li_2nd) > 1:
            if li_2nd[0].values() == li_2nd[1].values():
                random_index = random.choice(li_2nd).keys()[0]
                #print(random_index)
                if random_index in df_vote.index:
                    df_vote = df_vote.drop(random_index)
                    print(random_index, "is deleted")
    # Do the students put their 'student ID' correctly?
    #                           'Name'       correctly?
    return df_vote
if __name__ =='__main__':

    df_list = []
    #Check if some argument have been passed
    #pass the path of the excel file
    if(len(sys.argv) > 3):
        # sys.argv[1]: input excel file for students data within group informations
        # sys.argv[2]: the number of classes(if have 4 classes put '4')
        file_path = sys.argv[1]
        if(os.path.isfile(file_path)==False):
            print("Error: the file specified does not exist.")
        else:
            # read the input excel file
            df_original = pd.read_excel(file_path)
            for i in range(int(sys.argv[2])):
                df_class = pd.read_excel(file_path, i)
                if df_class.any : # have to fix later. but it works on 'EM_Fall_2017.xlsx'
                    print('Session '+ str(i) +'is opened correctly')
                df_list.append(df_class)
            input_path = sys.argv[3]
            if(os.path.isfile(input_path)==False):
                print("Error: the file specified does not exist")
            df_vote = pd.read_excel(input_path)
            #print(df_vote)
    else:
        print("You have put the path of input excel file and the number of sheets it has, for example: \n python dataframing.py /home/input.xlsx 4 ./voting_data.xlsx")
        #return
    #Check the errors and remove it before append to student roster.
    df_vote = ErrorCheck(df_vote)

    ## make a new series for append to the dataframe finally
    #Sample one of the row and matching the students in df_vote(dataframe of voting values)
    #                                                to df_list(dataframe of student roster)
    first_vote = pd.DataFrame()
    second_vote = pd.DataFrame()
    print(sys.argv[3])
    if str(sys.argv[3]).startswith('Vote_41'):
        print('41')
        df_4n = df_list[0]
    elif str(sys.argv[3]).startswith('Vote_42'):
        print('42')
        df_4n = df_list[1]
    elif str(sys.argv[3]).startswith('Vote_43'):
        print('43')
        df_4n = df_list[2]
    elif str(sys.argv[3]).startswith('Vote_44'):
        print('44')
        df_4n = df_list[3]

    #print(df_4n)
    for i in range(df_vote.shape[0]):
        # sample one of the row and get student ID of it
        sample = df_vote.sample(n=1)
        student_num = sample[u"Student ID No. (학번)"].values[0]
        sample = sample.dropna(axis = 1)

        idx = df_4n.index[df_4n['sid'] == student_num]
        #print(idx)
        if sample.columns.values[-1] == u"Let's vote.":
            temp1 = pd.DataFrame(sample[u"Let's vote."].values[0], index = idx, columns= ['1st_QLA_'])
            #print(temp1)
            first_vote = first_vote.append(temp1)
        if sample.columns.values[-1] == u"Let's vote again.":
            temp2 = pd.DataFrame(sample[u"Let's vote again."].values[0], index = idx, columns=['2nd_QLA_15'])
            #print(temp2)
            second_vote = second_vote.append(temp2)
            if sample.columns.values[-2] == u"Let's vote.":
                temp2_1 = pd.DataFrame(sample[u"Let's vote."].values[0], index = idx, columns= ['1st_QLA_15'])
                #print(temp2_1)
                first_vote = first_vote.append(temp2_1)
        #print(temp1)
        df_vote = df_vote.drop(sample.index)
    #print(first_vote)
    # append the new colums of '1st_QLA_15', '2nd_QLA_15'
    book = load_workbook(sys.argv[1])
    writer = pd.ExcelWriter(sys.argv[1], engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    result = pd.concat([df_4n, first_vote], axis = 1)
    result = pd.concat([result, second_vote], axis = 1)

    if sys.argv[3].startswith('Vote_41'):
        result.to_excel(writer,"EM_41")
    elif sys.argv[3].startswith('Vote_42'):
        result.to_excel(writer,"EM_42")
    elif sys.argv[3].startswith('Vote_43'):
        result.to_excel(writer,"EM_43")
    elif sys.argv[3].startswith('Vote_44'):
        result.to_excel(writer,"EM_44")
    writer.save()
    #print(result)

    # matching the answer by student ID.
