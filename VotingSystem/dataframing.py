# -*- coding; utf-8 -*-
# SKKU 2017
#
# This is the script for saving the voting results of in-class activity at dataframe.
# It is optimized for the class GEDB005-41, in fall semester of 2017 in SKKU
# This dataframe is for input data of analysing communication of students in same group during class activities.

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys
import os
import xlrd



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
                df_list.append(df_original)
            input_path = sys.argv[3]
            if(os.path.isfile(input_path)==False):
                print("Error: the file specified does not exist")
            df_vote = pd.read_excel(input_path)
            #print(df_vote)
    else:
        print("You have put the path of input excel file and the number of sheets it has, for example: \n python dataframing.py /home/input.xlsx 4 ./voting_data.xlsx")
        #return

    ## checking the mistaken of inputs
    # Do the students put their 'student ID' correctly?
    #                           'Name'       correctly?
    # Does any student put the answer twice at 1st and 2nd voting?


    ##Select one of the line and matching the students in df_vote to df_list
    first_vote = pd.DataFrame()
    second_vote = pd.DataFrame()

    for i in range(df_vote.shape[0]):
        sample = df_vote.sample(n=1)
        sample = sample.dropna(axis = 1)
        #print(sample)
        student_num = sample["Student ID No. (학번)"].values[0]

        #print(sample["Let's vote."].values)
        df_41 = df_list[0]
        idx = df_41.index[df_41['sid'] == student_num]
        #print(index)
        #print(sample.columns.values)
        if sample.columns.values[-1] == "Let's vote.":
            temp1 = pd.DataFrame(sample["Let's vote."].values[0], index = idx, columns= ['1st_QLA_15'])
            first_vote = first_vote.append(temp1)
        if sample.columns.values[-1] == "Let's vote again.":
            temp2 = pd.DataFrame(sample["Let's vote again."].values[0], index = idx, columns=['2nd_QLA_15'])
            second_vote = second_vote.append(temp2)
        #print(temp1)
        df_vote = df_vote.drop(sample.index)
    print(first_vote)
    # append the new colums of '1st_QLA_15', '2nd_QLA_15'
    # make a new series for append to the dataframe finally
    df_41 = df_41.merge(first_vote, left_index = True, right_index = True )
    df_41 = df_41.merge(second_vote, left_index = True, right_index = True )
    print(df_41)
    # matching the answer by student ID.
