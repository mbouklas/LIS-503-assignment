import openpyxl
import excel
# Imports the excel.py file that contains some of the functions used
from openpyxl import load_workbook
workbook = load_workbook(filename = "studentdata.xlsx")
# Loads the openpyxl module and loads the Excel workbook.

def get_int_input():
    sOption_Number_Valid = False
    iOption_Number = None
    while sOption_Number_Valid == False:
        try:
            sOption_Number_Text = input('Please choose an option:')
            iOption_Number = int(sOption_Number_Text)
            sOption_Number_Valid = True
        except ValueError:
            print('Invalid Number')
    return iOption_Number


def choose_option():
    
    print('''Please select an option:

    1. Average Grade by Student
    2. Missing Assignments by Student
    3. Missing Assignments by Assignment
    ''')

    iOption_Number = get_int_input()
    #Prompts the user to retry if they use text instead of an integer.
    while iOption_Number < 1 or iOption_Number > 3:
        print('Invalid Number')
        iOption_Number = get_int_input()
        print("Got user option", iOption_Number)
    return iOption_Number
# Prompts user to retry if they use invalid numbers.


def get_average_by_student(name):
    #Get grades for student
    lStudent_grades = excel.get_grades_by_name(name)
    #Get weightings for assignment types
    dWeightings = excel.load_weightings()
    #Compute average by type
    #Compute average for all assignments
    lWeighted_grades = []
    for assignment_dict in lStudent_grades:
        #Get category name from assignment dict
        sCategory = assignment_dict["Category"]
        #Get category weight from weightings dictionary
        fWeight = dWeightings[sCategory]

        if assignment_dict["Grade"] == "M":
            iGrade = 0.0
            lWeighted_grades.append(iGrade)
        elif assignment_dict["Grade"] != "X":
            iGrade = int(assignment_dict["Grade"])
            lWeighted_grades.append(sum(iGrade)/len(iGrade))
    fAvg = sum(lWeighted_grades)/len(lWeighted_grades)
    #This is not how you calculate a weighted average, fix
    print(lWeighted_grades)
    return fAvg       

def missing_assign_student():
    lStudent_grades = excel.get_student_names(name)
    lMissing_Assignments = []
    for assignment_dict in lStudent_grades:
        if assignment_dict["Grade"] == "M":
            lMissing_Assignments.append(sAssignment_name.value)
            
    return print(lMissing_Assignments)


def get_string_prompt(category):
    return input('Please enter {} name: '.format(category))


                  
def option_result():
    iOption_Number = choose_option()
    
    if iOption_Number == 1:
        print('Average Grade by Student')
        sStudent_Name_Average = get_string_prompt("student")
        fAvg = get_average_by_student(sStudent_Name_Average)
        print("{} average is: {}".format(sStudent_Name_Average, fAvg))

    if iOption_Number == 2:
        print('Missing Assignments by Student')
        sStudent_Name_Missing = get_string_prompt("student")
        sMA = missing_assign_student
        print('{} is missing the following assignments: {}'.format(sStudent_Name_Missing, sMA)) 

    if iOption_Number == 3:
        print('Missing Assignments by Assignment')
        sAssignment_Name = get_string_prompt("assignment")


while True:
    option_result()
