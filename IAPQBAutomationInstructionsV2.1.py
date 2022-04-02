# -*- coding: utf-8 -*-
"""
Spyder Editor
This is a script for converting Question Bank of TXT format into XLSX format
Specially applicable for Question Uploading in IAP (supports AIKEN Format)
Version 2.1
"""

import xlsxwriter # import XLSX writer package
import re # import Regex package

class Question:

    """ Question Attributes """
    question_number=0
    title=""
    problem_statement=""
    optionA=""
    optionB=""
    optionC=""
    optionD=""
    optionE=""
    optionF=""
    solution=""
    maximum_mark=""
    topic=""
    tag=""
    sharee=""
    filename=""
    attributes_list = []
    error_code=""

    def __init__(self,individual_questions,topic_name_input,tag_name_input,sharee_name_input):
        Question.question_number+=1
        self.attributes_list = individual_questions
        self.maximum_mark = 1 # default value equals to '1'
        self.topic = topic_name_input
        self.tag = tag_name_input
        self.sharee = sharee_name_input
        self.filename = "format" # default XLSX filename is 'format'
        self.title = self.tag + str(Question.question_number)

    def attributes_allocator(question_object,individual_questions):

        if len(individual_questions)<4: # If any combination of Problem Statement, Option A, Option B and ANSWER is not supplied
            question_object = None

        elif len(individual_questions)>8: # If more than 6 choices are supplied
            question_object = None

        else:
            for item in individual_questions: # For correct inputs
                if re.match("^A. ",item) or re.match("^A\) ",item):
                    question_object.optionA = item[3:]
                elif re.match("^B. ",item) or re.match("^B\) ",item):
                    question_object.optionB = item[3:]
                elif re.match("^C. ",item) or re.match("^C\) ",item):
                    question_object.optionC = item[3:]
                elif re.match("^D. ",item) or re.match("^D\) ",item):
                    question_object.optionD = item[3:]
                elif re.match("^E. ",item) or re.match("^E\) ",item):
                    question_object.optionE = item[3:]
                elif re.match("^F. ",item) or re.match("^F\) ",item):
                    question_object.optionF = item[3:]
                elif re.match("^ANSWER. ",item):
                    question_object.solution = item[8:]
                else:
                    question_object.problem_statement = item

""" Excel Sheet Management Module """
def excel_writer(quesiton_bank_with_zero_error):
    OutWorkBook = xlsxwriter.Workbook("format.xlsx")
    OutSheet = OutWorkBook.add_worksheet()

    OutSheet.write("A1","Title")
    OutSheet.write("B1","Problem Statement")
    OutSheet.write("C1","Option A")
    OutSheet.write("D1","Option B")
    OutSheet.write("E1","Option C")
    OutSheet.write("F1","Option D")
    OutSheet.write("G1","Option E")
    OutSheet.write("H1","Option F")
    OutSheet.write("I1","Solutions")
    OutSheet.write("J1","Max_Marks")
    OutSheet.write("K1","Topic")
    OutSheet.write("L1","Tags")
    OutSheet.write("M1","Sharee")
    OutSheet.write("N1","Filename")

    for item in range(len(quesiton_bank_with_zero_error)):
        OutSheet.write(item+1,0,(quesiton_bank_with_zero_error[item]).title)
        OutSheet.write(item+1,1,(quesiton_bank_with_zero_error[item]).problem_statement)
        OutSheet.write(item+1,2,(quesiton_bank_with_zero_error[item]).optionA)
        OutSheet.write(item+1,3,(quesiton_bank_with_zero_error[item]).optionB)
        OutSheet.write(item+1,4,(quesiton_bank_with_zero_error[item]).optionC)
        OutSheet.write(item+1,5,(quesiton_bank_with_zero_error[item]).optionD)
        OutSheet.write(item+1,6,(quesiton_bank_with_zero_error[item]).optionE)
        OutSheet.write(item+1,7,(quesiton_bank_with_zero_error[item]).optionF)
        OutSheet.write(item+1,8,(quesiton_bank_with_zero_error[item]).solution)
        OutSheet.write(item+1,9,(quesiton_bank_with_zero_error[item]).maximum_mark)
        OutSheet.write(item+1,10,(quesiton_bank_with_zero_error[item]).topic)
        OutSheet.write(item+1,11,(quesiton_bank_with_zero_error[item]).tag)
        OutSheet.write(item+1,12,(quesiton_bank_with_zero_error[item]).sharee)
        OutSheet.write(item+1,13,(quesiton_bank_with_zero_error[item]).filename)

    OutWorkBook.close()

    print("Please check your project folder. Your Excel file is generated")

""" Filter the work process """
def handle_error(quesiton_bank_with_zero_error, question_bank_error_positions):
    if not question_bank_error_positions:
        excel_writer(quesiton_bank_with_zero_error)
    else:
        print("The Question Bank has error on the following questions\n")
        for item in question_bank_error_positions:
            print("Question "+str(item))
        print("Please Try Again after checking the formats of the mentioned questions")
        exit()

""" Question Bank error identifier (if any) """
def question_bank_rectifier(quesiton_bank_with_probable_error):
    error_positions = []
    correct_questions = []
    for position in range(len(quesiton_bank_with_probable_error)):
        if quesiton_bank_with_probable_error[position]==None:
            error_position = position+1
            error_positions.append("Question "+str(error_position))
        else:
            correct_questions.append(quesiton_bank_with_probable_error[position])
    return correct_questions,error_positions

""" Instantiate the Questions with proper attributes """
def question_bank_creator(individual_questions_list,topic_name_input,tag_name_input,sharee_name_input):
    quesiton_objects_list = []
    for individual_questions in individual_questions_list:
        question_object = Question(individual_questions,topic_name_input,tag_name_input,sharee_name_input)
        Question.attributes_allocator(question_object,individual_questions)
        quesiton_objects_list.append(question_object)
    return quesiton_objects_list

""" Distribute the Questions """
def question_seperator(docs_string):
    question_bank_list = []
    single_question = []
    for item in range(len(docs_string)):
        if re.match("^ANSWER. ",docs_string[item]):
            single_question.append(docs_string[item])
            question_bank_list.append(single_question)
            single_question = []
        else:
            single_question.append(docs_string[item])
    return question_bank_list

"""Reading the Text File"""
def question_reader(question_bank):
    question_bank_file = open(question_bank,"r+")
    question_bank_converted_to_lines_of_string = question_bank_file.read().splitlines()
    question_bank_file.close()
    return question_bank_converted_to_lines_of_string

def main():
    """Collect User Inputs"""
    file_to_read_input=str(input("Enter Name of Your Text File. The file should be inside your project folder\n:>"))
    topic_name_input=str(input("Enter Topic Name\n:>"))
    tag_name_input=str(input("Enter Tag Name\n:>"))
    sharee_name_input=str(input("Enter Name of the Sharees comma seperated, without @infosys.com\n:>"))

    docs_string = question_reader(file_to_read_input+".txt") # Reads the .txt file
    individual_questions_list = question_seperator(docs_string) # Separates strings to individual questions
    quesiton_bank_with_probable_error = question_bank_creator(individual_questions_list,topic_name_input,tag_name_input,sharee_name_input) # Creates Question_Bank with Question Objects
    quesiton_bank_with_zero_error, question_bank_error_positions = question_bank_rectifier(quesiton_bank_with_probable_error)
    handle_error(quesiton_bank_with_zero_error, question_bank_error_positions)

if __name__ == "__main__":
    main()