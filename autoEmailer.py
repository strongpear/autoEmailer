# Author: Jason Lee

from classes import *      # Class definitions
import csv                 # csv_reader
import configparser    # For first time configuration
import pandas as pd    # For reading .xlsx files
import os              # os.path to save to sentemails folder
import simplegmail     # For emailing and authentication

VERSION = 0.2

def main():
    # Print welcome banner
    printBanner(40)

    # First time setup config.
    firstTimeSetup()

    # Ask teacher what class this is for
    course = getCourse()
    print("You have chosen the course: " + course + ".")

    studentDataFileName = input("Student Data Filename (.csv file): ")
    # Get student Data

    # Input student data into a dictionary.
    # This dictionary is formatted like {student_ID : Student()}
    # So, the key is the student ID.
    roster = inputStudentData(studentDataFileName)

    # Get Gradebook data
    gradebookDataFileName = input("Gradebook Data Filename (.xlsx file): ")

    # Transforms the excel file into the csv format we want.
    transformGrades(gradebookDataFileName)
    with open("grades.csv", "r") as gradebookData:

        # TODO: Get URL's from user
        getURLs(gradebookData, course)

    with open("grades.csv", "r") as gradebookData:
        # Update if student is failing and failing assignments for those students
        getFailing(gradebookData, roster, course)

    # Confirm list of failing students with the user
    comfirmFailing(roster)

    # Send emails to parents
    sendEmails(roster, course)



# This function prints the inital welcome banner
def printBanner(width):
    print('*' * width)
    print("Parent Emailer".center(width))
    print(("Version " + str(VERSION)).center(width))
    print('*' * width, "\n")

# This is the first time setup
def firstTimeSetup():
    config = configparser.ConfigParser()
    config.read("emailconfig.ini")          # Read config file
    firstTime = config.get("personal", "first_time")
    if firstTime == "True":
        print("It seems that this is your first time using this.")
        print("If you haven't already, you NEED to read README.txt before using this.\n")
        while True:
            try:
                file = open("gmail_token.json", "r")
                file.close()
                break
            except FileNotFoundError:
                print("You need to login to Gmail first. This is required to send emails.")
                input("Press Enter to continue to login to Gmail. ")
                gmail = simplegmail.Gmail()

        print("Please answer some questions for first time setup.")
        # Get teacher name
        while True:
            name = input("What's your first and last name? (ex. Jason Lee) ").split()
            if len(name) == 2:
                break
            print("Please type only your first and last name separated by a space. ")

        # Set config values
        config.set("personal", "last_name", name[1])
        config.set("personal", "first_name", name[0])

        # Get teacher email
        while True:
            email = input("What is your work email address? ")
            if email.find("@") != -1:
                break
            print("Please enter a valid email address. ")
        config.set("personal", "email", email)

        # Get classes taught
        while True:
            classes = input("What courses do you teach? (ex. Algebra 1, Pre-AP Algebra 2) ").split(', ')
            if len(classes) > 0:
                break
            print("Please enter in at least one course that you teach.")

        # Reset Courses
        config["courses"] = {}

        # Set courses
        for i in range(len(classes)):
            config.set("courses", "course " + str(i + 1), classes[i])

        # Write all info into config file
        with open("emailconfig.ini", "w") as file:
            config.set("personal", "first_time", "False")
            config.write(file)
        print("\nConfiguration successful. If you want to change these settings, open emailconfig.ini")
        print("and change the value of first_time to True\n")

# This function gets which class the teacher is sending to.
def getCourse():

    # Read courses from config file
    config = configparser.ConfigParser()
    config.read("emailconfig.ini")

    # List comprehension of courses
    courses = [option for option in config['courses'].values()]

    # Print menu for courses
    for i in range(len(courses)):
        print(i + 1, ": " + courses[i])
    while True:
        choice = input("What course is this for? Pick a number.\n")

        # Input validation
        try:
            if int(choice) > 0 and int(choice) <= len(courses):
                break
            print("Please pick a valid number. ")
        except (ValueError):
            print("Please type in a number. ")

    # Return name of course
    return courses[int(choice) - 1]

# This function returns a dictionary {student_ID : Student()}
def inputStudentData(fileName):

    while True:
        try:
            if fileName.find('.csv') != -1:
                file = open(fileName, "r")
                break
            else:
                print("Invalid file name. Did you type it incorrectly? Is it in the same folder as this program?\n")
                fileName = input("Student Data Filename (.csv file): ")
        except FileNotFoundError:
            print("Invalid file name. Did you type it incorrectly? Is it in the same folder as this program?\n")
            fileName = input("Student Data Filename (.csv file): ")

    # Create dictionary
    students = dict()
    reader = csv.reader(file, delimiter = ',')

    # Use this to skip the first line
    count = 0
    for row in reader:
        if count == 0:
            if row[0] == "Timestamp" and row[3] == "Student ID" and row[7] == "Parent Email":
                count += 1
            else:
                print("Your student data is not formatted correctly. Use the google form in README.txt.")
                exit(2)
        else:
            temp = Student()

            # Person(Last Name, First Name, Email)
            temp.person = Person(row[1], row[2], row[4])
            temp.parent = Person(row[5], row[6], row[7])

            # row[3] is their student ID
            students[row[3]] = temp
            count += 1
    print("Successfully loaded", count - 1, "students.")
    file.close()
    return students

# This function transforms the xlsx file that the gradebook gives us into
# a csv file that I know how to deal with.
def transformGrades(fileName):

    # Read excel file
    while True:
        try:
            read_file = pd.read_excel(fileName)
            break
        except (FileNotFoundError, ValueError):
            print("Invalid file name. Did you type it incorrectly? Is it in the same folder as this program?\n")
            fileName = input("Gradebook data Filename (.xlsx file): ")
    print("Converting file to .csv...\n")

    # Convert to csv
    read_file.to_csv("grades.csv", index = None, header = True)

    lines = list()
    print("Fixing formatting...\n")
    with open("grades.csv", "r") as file:

        # Read csv file
        reader = csv.reader(file)
        for row in reader:
            lines.append(row)

    # The gradebook's first 9 lines are header lines. We delete those.
    del lines[0:8]

    count = 0

    # Make sure Gradebook data is formatted correctly.
    if lines[0][2] != "Grade":
        print("It seems like the format of the gradebook is incorrect.")
        print("It's likely that you may have forgotten to check 'Show Grades' when creating the grade report.")
        exit(3)
    # The second line contains the titles. We are fixing it so only the
    # title of the assignment is left.
    for item in lines[0]:
        if count > 3:               # Assignments start on column 4.
            item = item.split("\n") # It's formatted with 5 lines, we only want the last one.
            lines[0][count] = item[-1]
        count += 1

    # Overwrite the file using our new lines that we created.
    with open("grades.csv", "w", newline = '') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerows(lines)

# This function sets up URLs.ini for the user to input URLs.
def getURLs(file, course):
    # Open gradebook
    reader = csv.reader(file, delimiter = ',')
    count = 0
    URLs = {}

    # Add to URLs Dictionary, just in case I want to add manual adding onto here.
    for row in reader:
        if count == 0:
            for i in range(4, len(row)):
                URLs[row[i]] = ""
            count += 1
        else:
            break

    # Open Config file
    config = configparser.ConfigParser()
    config.read("URLs.ini")
    courseURLs = course + " URLs"
    if courseURLs not in config:
        config[courseURLs] = {}

    if config[courseURLs] != {}:
        choice = input("Are you using the same " + course + " URL's as before? If this is a new 6 weeks, then probably not. (Y/N)")
        if choice.upper() in ["Y", "YES"]:
            input("Okay, great. Press Enter to Continue.")
            return
    # Reset config file
    config[course + " URLs"] = {}

    # Set config file with titles of assignments
    for item in URLs:
        config.set(courseURLs, item, URLs[item])
    with open("URLs.ini", "w") as file:
        config.write(file)

    # Prompt user to update URLs.ini
    print("If you want to link any URLs in the email, please go to URLs.ini and input those in.")
    input("Press Enter after you have finished...")
    input("Are you sure you have finished? Press Enter again.\n")

# This function updates students failure status.
def getFailing(file, roster, course):

    # Read file
    reader = csv.reader(file, delimiter = ',')
    config = configparser.ConfigParser()
    config.read("URLs.ini")
    # Use to count the rows
    count = 0

    # Local variables to help
    studentIDs = []
    URLs = ['','','','']     # This is mega jank, but it works for now.
    titles = []

    # Update URLs from Config file
    courseURLs = course + " URLs"
    for link in config[courseURLs].values():
        URLs.append(link)

    for row in reader:

        # First row has the titles
        if count == 0:
            for i in range(len(row)):
                titles.append(row[i])
            count += 1

        # Rest of the rows contain grade data
        else:
            # Check if failing
            if checkFailing(int(row[2])) == True:
                failingStudentID = row[0].strip()
                try:
                    check = roster[failingStudentID]
                    studentIDs.append(failingStudentID)
                except KeyError:
                    print("***Student with ID", failingStudentID, "was not found in roster. Type N.***")
                    continue

                # Update class variable to True if failing.
                roster[failingStudentID].isFailing = True
                for i in range(4, len(row)):
                    # If X or blank, ignore.
                    excludedValues = ['X', '']
                    if row[i].strip() in excludedValues:
                        continue
                    if row[i].strip() == 'M' or checkFailing(float(row[i])) == True:
                        temp = Assignment(titles[i], row[i], URLs[i])
                        roster[failingStudentID].failingAssignments.append(temp)
            count += 1
    return studentIDs

# This function checks with the user that the info of their failing students is correct.
def comfirmFailing(roster):
    print("\nList of failing students, and their parent's emails:")
    for student in roster:
        if roster[student].isFailing == True:
            print(roster[student].person.firstName, roster[student].person.lastName + ",", roster[student].parent.email)
    response = input("\nDoes this look correct? (Y/N)\n")
    if response.lower() == "y" or response.lower() == "yes":
        print("\nGreat, sending emails...\n")
    else:
        print("Okay, something went wrong then. If it's emails then change them in the Student Data file.")
        print("If you have ID ###### was not found in roster, update your student data file with their information on a new row.")
        print("If it's something else, then that is not good and don't use this unless you can figure it out.")
        input("Press Enter to Exit.")
        exit(4)

# This function sends out the emails to students and parents.
def sendEmails(roster, course):

    # Initial info to user
    print("Before we send emails, the file emailtemplate.txt houses the template of the email. ")
    print("Make sure you read the README.txt, especially if you want to change the template. ")
    print("You can see an example email by opening exampleemail.txt.")
    input("Press Enter to continue\n")

    # Get teacher's email
    config = configparser.ConfigParser()
    config.read("emailconfig.ini")
    teacherEmail = config.get("personal", "email")

    # Open Gmail
    gmail = simplegmail.Gmail()

    # Only email failing students.
    count = 0
    for student in roster:
        if roster[student].isFailing == True:

            # Create list of failing assignments in string form
            assignments = "Title, Grade, URL (If available)\n"
            for item in roster[student].failingAssignments:
                assignments = assignments + item.title + ", " + item.grade + ", " + item.URL + "\n"

            # Create dictionary to swap keywords in emailbody.txt
            keywords = {"PARENT_FIRST_NAME" : roster[student].parent.firstName,
                        "PARENT_LAST_NAME" : roster[student].parent.lastName,
                        "STUDENT_FIRST_NAME" : roster[student].person.firstName,
                        "STUDENT_LAST_NAME" : roster[student].person.lastName,
                        "COURSE_NAME" : course,
                        "FAILING_ASSIGNMENTS" : assignments,
                        "TEACHER_EMAIL" : teacherEmail}
            # Open outline
            with open("emailtemplate.txt", "r") as fileOutline:
                text = fileOutline.read()

                # Replace keywords in outline
                for i, j in keywords.items():
                    text = text.replace(i, j)

            # Generate output file so user can see sent mail.
            fileName = "sentEmail" + roster[student].person.lastName + ".txt"
            with open(os.path.join("sentemails", fileName), "w") as fileBody:
                fileBody.write(text)
            with open(os.path.join("sentemails", fileName), "r") as emailBody:
                body = emailBody.read()

                # Set email parameters
                params = { "to" : roster[student].parent.email,
                           "sender" : "vrhsmath.no.reply@gmail.com",
                           "subject" : "Your " + course + " Student",
                           "msg_html" : "<pre style = 'font-family: Arial'>" + text + "</pre>"}
                print("Sending email to", roster[student].parent.firstName, roster[student].parent.lastName, "...")

                # Send email
                message = gmail.send_message(**params)
                count += 1
    print("Successfully sent", count, "emails.")
    input("Press Enter to Exit.")
# Short function for readability
def checkFailing(grade):
    if grade < 70:
        return True
    else:
        return False
main()
