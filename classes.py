class Assignment:
    def __init__(self, title, grade, URL):
        self.title = title
        self.grade = grade
        self.URL = URL
class Person:
    def __init__(self, lastName, firstName, email):
        self.lastName = lastName
        self.firstName = firstName
        self.email = email
class Student:
    person = Person("default", "default", "default")
    parent = Person("default", "default", "default")
    isFailing = False
    def __init__ (self):
        self.failingAssignments = []
    def addAssignment(s):
        failingAssignments.append(s)
