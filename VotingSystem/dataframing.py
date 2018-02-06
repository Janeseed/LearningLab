# -*- coding: utf-8 -*-

# SKKU 2017
#
# This is the script for saving the voting results of in-class activity at dataframe.
# It is optimized for the class GEDB005-41, in fall semester of 2017 in SKKU
# This dataframe is for input data of analysing communication of students in same group during class activities.

import gc
import os
import random
import sys
import numpy as np
import pandas as pd
from functools import reduce


def ErrorCheck(df_vote):
    ## checking the mistaken of inputs
    # Does any student put the answer twice at 1st and 2nd voting?
    df_new = df_vote
    df_sample = df_new.groupby(u"Student ID No. (학번)").groups

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
            # get the value of first and second vote.
            row = df_new.ix[li_index[i]]
            first_vote = row["Let's vote."]
            second_vote = row["Let's vote again."]

            # split two answered row into two
            if (not np.isnan(first_vote)) and (not np.isnan(second_vote)):
                df_new = df_new.drop(li_index[i])
                new_row1 = row
                new_row1.drop(["Let's vote."])
                df_new = df_new.append(new_row1)

                new_row2 = row
                new_row2.drop(["Let's vote again."])
                new_row2.name = 1004
                df_new = df_new.append(new_row2)

                li_1st.append({1004: new_row2["Let's vote."]})
                li_2nd.append({li_index[i]: new_row1["Let's vote again."]})

            if not np.isnan(first_vote) and np.isnan(second_vote):
                li_1st.append({li_index[i]: first_vote})
            if not np.isnan(second_vote) and np.isnan(first_vote):
                li_2nd.append({li_index[i]: second_vote})
        print(li_1st, li_2nd)
            #the print results are like below
            #example1. ([{13: 3.0}], [{128: 3.0}, {129: 3.0}])
            #example2. ([{5: 4.0}, {11: 4.0}], [{5: 4.0}, {92: 4.0}])
            #example3. ([{5: 4.0}, {82: 2.0}, {61: 2.0}])

            #If it has double voting data and the value is same, deleted
        while len(li_1st) > 1:
            print(li_1st)
            dic_row = {}
            # save every duplicated data to dictionary and delete those data at dataframe
            for j in range(len(li_1st)):
                index = list(li_1st[j].keys())[0]
                row = df_new.xs(index)
                dic_row.update({index: row})
                df_new = df_new.drop(index)

            # Using reduce function, choose one of the duplicated data
            li_1st = reduce(lambda x, y: x if list(x.values())[0] == list(y.values())[0] else (y if list(x.keys())[0] > list(y.keys())[0] else x), li_1st)
            print(li_1st)
            print("delete duplicated data")

            #append finally choosed data to dataframe
            final_index = list(li_1st.keys())[0] # index that selected by the reduce fn. e.g. if li_1st is {5: 4.0}, final_index = 5
            row_to_append = dic_row.get(final_index)
            df_new = df_new.append(row_to_append, verify_integrity=True)


        # Do the same-thing to second vote
        while len(li_2nd) > 1:
            print(li_2nd)
            dic_row = {}
            # save every duplicated data to dictionary and delete those data at dataframe
            for j in range(len(li_2nd)):
                index = list(li_2nd[j].keys())[0]
                row = df_new.xs(index)
                dic_row.update({index: row})
                df_new = df_new.drop(index)

            # Using reduce function, choose one of the duplicated data
            li_2nd = reduce(lambda x, y: x if list(x.values())[0] == list(y.values())[0] else (y if list(x.keys())[0] > list(y.keys())[0] else x), li_2nd)
            print(li_2nd)
            print("delete duplicated data")
            #append finally choosed data to dataframe
            final_index = list(li_2nd.keys())[0] # index that selected by the reduce fn. e.g. if li_2nd is {5: 4.0}, final_index = 5
            row_to_append = dic_row.get(final_index)
            df_new = df_new.append(row_to_append, verify_integrity=True)
    #print(df_new)
    return df_new

if __name__ =='__main__':

    df_list = []
    #Check if some argument have been passed
    #pass the path of the excel file
    if(len(sys.argv) > 2):
        # sys.argv[1]: input excel file for students data within group information
        # sys.argv[2]: directory name of data stacked
        file_path = sys.argv[1]
        if(os.path.isfile(file_path)==False):
            print("Error: the file specified does not exist.")
        else:
            # read the input excel file
            df_original = pd.ExcelFile(file_path)
            for i in range(len(df_original.sheet_names)):
                df_class = pd.read_excel(file_path, i)
                if df_class.any : # have to fix later. but it works on 'EM_Fall_2017.xlsx'
                    print('Session '+ str(i) +'is opened correctly')
                df_list.append(df_class)
            dirname = sys.argv[2]
            if(os.path.isdir(dirname)==False):
                print("Error: the directroy name specified does not exist")

    else:
        print("You have put the path of input excel file and the number of sheets it has, for example: \n python dataframing.py /home/input.xlsx /Documents/learning_lab/votingdata")
        #return
    # Do the saving For the all files in directories
    dirname = sys.argv[2]
    filenames = os.listdir(dirname)
    for filename in filenames:
        print(filename)
        full_filename = os.path.join(dirname, filename)

        df_vote = pd.read_excel(full_filename)
        #print(df_vote)
        # Check the errors and remove it before append to student roster.
        df_vote = ErrorCheck(df_vote)

        ## make a new series for append to the dataframe finally
        # Sample one of the row and matching the students in df_vote(dataframe of voting values)
        #                                                to df_list(dataframe of student roster)
        # Find out the class_num and lecture_num from filename
        filename = filename.replace(" ", "_")
        info_string = filename.split("_")[1]
        class_num = info_string.split("-")[0] # e.g. 41
        section_tag = info_string.split("-")[1] # e.g. PDE, LA, QLA
        lecture_num = info_string.split("-")[2] # e.g. 15

        if class_num == "41":
            df_4n = df_list[0]
        elif class_num == "42":
            df_4n = df_list[1]
        elif class_num == "43":
            df_4n = df_list[2]
        elif class_num == "44":
            df_4n = df_list[3]

        # print(df_4n)
        first_vote = pd.DataFrame()
        second_vote = pd.DataFrame()
        for i in range(df_vote.shape[0]):
            # sample one of the row and get student ID of it
            sample = df_vote.sample(n=1)
            student_num = sample[u"Student ID No. (학번)"].values[0]
            sample = sample.dropna(axis = 1)

            idx = df_4n.index[df_4n['sid'] == student_num]

            # Do the students put their 'student ID' correctly?
            if not idx.any:
                print("Student ID is not found")

            # print(idx)
            if sample.columns.values[-1] == u"Let's vote.":
                temp1 = pd.DataFrame(sample[u"Let's vote."].values[0], index = idx, columns= ['1st_'+ str(lecture_num) + "_" + str(section_tag)])
                #print(temp1)
                first_vote = first_vote.append(temp1, verify_integrity=True)
            if sample.columns.values[-1] == u"Let's vote again.":
                temp2 = pd.DataFrame(sample[u"Let's vote again."].values[0], index = idx, columns=['2nd_' + str(lecture_num) + "_" + str(section_tag)])
                #print(temp2)
                second_vote = second_vote.append(temp2, verify_integrity=True)
                if sample.columns.values[-2] == u"Let's vote.":
                    temp2_1 = pd.DataFrame(sample[u"Let's vote."].values[0], index = idx, columns= ['1st_'+ str(lecture_num) + "_" + str(section_tag)])
                    #print(temp2_1)
                    first_vote = first_vote.append(temp2_1, verify_integrity=True)
            # print(temp1)
            df_vote = df_vote.drop(sample.index)
        print("first vote", first_vote)
        print("second vote", second_vote)
        # append the new columns of '1st_QLA_15', '2nd_QLA_15'
        result = df_4n.join(first_vote, how = "outer")
        print("result", result)
        if second_vote.any:
            result = result.join(second_vote, how = "outer")
            print("result", result)

        if class_num == "41":
            df_list[0] = result
        elif class_num == "42":
            df_list[1] = result
        elif class_num == "43":
            df_list[2] = result
        elif class_num == "44":
            df_list[3] = result

        del first_vote # remove the context of first_vote
        del second_vote # remove the context of first_vote
        del df_vote

        gc.collect()

    #print(df_list[0])
    writer = pd.ExcelWriter("Voting_data.xlsx", engine='openpyxl')
    df_list[0].to_excel(writer, "EM_41")
    df_list[1].to_excel(writer, "EM_42")
    df_list[2].to_excel(writer, "EM_43")
    df_list[3].to_excel(writer, "EM_44")
    writer.save()
