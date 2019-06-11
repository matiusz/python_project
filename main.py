#sala tabelka zajętości






import random
import math
import unicodedata
from openpyxl import *
import collections

class Room:
    @staticmethod
    def DeadRoom():
        room = Room("", "")
        room.classes = []
        return room
    @staticmethod
    def findByString(roomString, RoomsList):
        roomStringList = roomString.split()
        if(len(roomStringList)<2):
            return Room.DeadRoom()
        for room in RoomsList:
            if(room.building==roomStringList[0]):
                if(room.number==roomStringList[1]):
                    return room
        #TODO remove prints
        #print("Room not found")
        return Room.DeadRoom()
    def __init__(self, Building, Number):
        if (Building != None):
            self.building = Building
            self.classes = []
        else:
            self.building = ""
        if (Number != None):
            self.number = Number
        else:
            self.number = ""
    def __str__(self):
        return(self.building+" "+self.number)
class Person:
    @staticmethod
    def findSimiliarNames(string, PersonList, n=5):
        suggestedNames = []
        for persons in PersonList:
            if(distance(string, persons.lastname+" "+persons.firstname) == 0 or distance(string, persons.firstname+" "+persons.lastname) == 0):
                print(persons.lastname+" "+persons.firstname)
                return [(1, persons)]
            if(distance(string, persons.lastname+" "+persons.firstname)<=n or distance(string, persons.firstname+" "+persons.lastname)<=n):
                suggestedNames.append(persons)
            if(distance(string, persons.lastname) <= 2): #
                suggestedNames.append(persons)
        if suggestedNames!=[]:
            #print("Did you mean:")
            for i, suggestion in enumerate(suggestedNames, 1):
                #print(str(i) + ": " + suggestion.lastname+" "+suggestion.firstname)
                return list(enumerate(suggestedNames, 1))

    @staticmethod
    def DeadPerson():
        person = Person("", "")
        person.degree = ""
        person.classes = []
        return person
    @staticmethod
    def findByName(LastName, FirstName, PersonList):
        for person in PersonList:
            if(LastName==person.lastname):
                if (FirstName==person.firstname):
                    return person
        #TODO remove prints
        #print("Person not found")
        Person.findSimiliarNames(LastName+" "+FirstName, PersonList)
        return Person.DeadPerson()
    def __init__(self, LastName, FirstName):
        self.firstname = FirstName
        self.lastname = LastName
        self.degree = ""
        self.classes = []
    def __str__(self):
        return(self.firstname+" "+self.lastname+", "+self.degree)
class Classes:
    def __init__(self, Subject):
        self.subject = Subject
    def __str__(self):
        return(self.subject+", "+str(self.person)+", "+str(self.classroom)+", "+self.day+" "+str(self.hourStart))
class Parser:
    @staticmethod
    def ParsePersons(persons):
        PersonList = []
        for row in persons.iter_rows(min_row=2):
            name = row[0].value.split()
            person = Person(name[0], name[1])
            PersonList.append(person)
            person.department = row[1].value
            person.job = row[2].value
            if (row[3].value!=None):
                person.degree = row[3].value
            person.additional = row[4].value
            person.pensum = row[5].value
            person.discount = row[6].value
            person.email = row[7].value
            person.room = row[8].value
            person.day = row[9].value
            person.time = row[10].value
        for x in PersonList:
            #TODO remove prints
            #print(x)
            pass
        return PersonList
    def ParseClassroom(classrooms):
        RoomsList = []
        for row in classrooms.iter_rows(min_row=2):
            if (row[0].value==row[1].value):
                continue
            classroom = Room(row[0].value, row[1].value)
            RoomsList.append(classroom)
            classroom.type = row[2].value
            classroom.capacity = row[3].value
            classroom.notes = row[4].value
        for x in RoomsList:
            #TODO remove prints
            #print(x)
            pass
        return RoomsList
    def ParseSchedule(PersonList, RoomsList, sheet):
        ClassesList = []
        for row in sheet.iter_rows(min_row=2):
            if(row[0].value=="---" or row[0].value==None or row[3].value == None):
                continue
            classes = Classes(row[3].value)
            ClassesList.append(classes)
            classes.course = row[0].value
            if(row[1]!=None):
                classes.semester = row[1].value
            else:
                classes.semester = ""
            if(row[2].value != None):
                classes.location = row[2].value
            else:
                classes.location = ""
            if (row[4].value != None):
                classes.chooseable = row[4].value
            else:
                classes.chooseable = row[5].value
            if (row[5].value!=None):
                classes.type = row[5].value
            else:
                classes.type = ""
            if (row[6].value != None):
                classes.group = row[6].value
            else:
                classes.group = ""
            if(row[7].value!=None):
                classes.timeframe = row[7].value
            else:
                classes.timeframe = ""
            if(row[9].value!=None):
                classes.faculty = row[9].value
            else:
                classes.faculty = ""
            nameString = row[10].value
            if (nameString!=None):
                name=nameString.split()
                classes.person = Person.findByName(name[0], name[1], PersonList)
                classes.person.classes.append(classes)
            else:
                classes.person = Person.DeadPerson()
            classroomString = row[11].value
            if (classroomString!=None):
                classes.classroom = Room.findByString(classroomString, RoomsList)
                classes.classroom.classes.append(classes)
            else:
                classes.classroom = Room.DeadRoom()
            if (row[12].value!=None):
                classes.week = row[12].value
            else:
                classes.week = ""
            if(row[12].value!=None):
                classes.week = row[12].value or ""
            else:
                classes.week
            classes.day = row[13].value or ""
            classes.hourStart = row[14].value or ""
            classes.hourEnd = row[15].value or ""
        for x in ClassesList:
            #TODO remove prints
            #print(x)
            pass
        return ClassesList
    @staticmethod
    def ParseAll(filename = '2019-2020(6) (1).xlsx'):
        workbook = load_workbook(filename)
        personsSheet = workbook["osoby"]
        roomsSheet = workbook["sale"]
        PersonsList = Parser.ParsePersons(personsSheet)
        RoomsList = Parser.ParseClassroom(roomsSheet)
        WinterSList = Parser.ParseSchedule(PersonsList, RoomsList, workbook["zima_s"])
        SummerSList = Parser.ParseSchedule(PersonsList, RoomsList, workbook["lato_s"])
        WinterNList = Parser.ParseSchedule(PersonsList, RoomsList, workbook["zima_n"])
        SummerNList = Parser.ParseSchedule(PersonsList, RoomsList, workbook["lato_n"])
        AllClasses = WinterSList+SummerSList+WinterNList+SummerNList

        return (PersonsList, RoomsList, AllClasses) # check if tuple order matches
        #Person.findSimiliarNames("Alda Witold", PersonsList, 5)

        
def printPersonData(person):
    print("\n" + str(person))
    print("Kontakt: " + person.email)
    print("Konsultacje: " + (person.room or "")+ " " + (person.day or "")+ " " + str(person.time or "") +"\n")
    print("Prowadzone zajęcia: ")
    classes_monday = []
    classes_tuesday = []
    classes_wednesday = []
    classes_thursday = []
    classes_friday = []
    classes_other = []
    '''for classes in person.classes:
        print(classes.subject + " " + str(classes.classroom) + " " + classes.day + " " + str(classes.hourStart))'''
    for classes in person.classes:
        if classes.day == "Pn":
            classes_monday.append(classes)
        elif classes.day == "Wt":
            classes_tuesday.append(classes)
        elif classes.day == "Sr":
            classes_wednesday.append(classes)
        elif classes.day == "Cz":
            classes_thursday.append(classes)
        elif classes.day == "Pt":
            classes_friday.append(classes)
        else:
            classes_other.append(classes)
    print("\nPoniedziałek:\n")
    if classes_monday == []:
            print("Brak zajęć")
    else:
        for classes in classes_monday:
            print(classes.subject + " " + str(classes.classroom) + " " + str(classes.hourStart))
    print("\nWtorek:\n")
    if classes_tuesday == []:
            print("Brak zajęć")
    else:
        for classes in classes_tuesday:
            print(classes.subject + " " + str(classes.classroom) + " " + str(classes.hourStart))
    print("\nŚroda:\n")
    if classes_wednesday == []:
            print("Brak zajęć")
    else:
        for classes in classes_wednesday:
            print(classes.subject + " " + str(classes.classroom) + " " + str(classes.hourStart))
    print("\nCzwartek:\n")
    if classes_thursday == []:
            print("Brak zajęć")
    else:
        for classes in classes_thursday:
            print(classes.subject + " " + str(classes.classroom) + " " + str(classes.hourStart))
    print("\nPiątek:\n")
    if classes_friday == []:
            print("Brak zajęć")
    else:
        for classes in classes_friday:
            print(classes.subject + " " + str(classes.classroom) + " " + str(classes.hourStart))
    print("\nPozostałe:\n")
    if classes_other == []:
            print("Brak zajęć" + "\n")
    else:
        for classes in classes_other:
            print(classes.subject + " " + str(classes.classroom) + " " + classes.day + " " + str(classes.hourStart))

# choose person from list of suggestions and call printPersonData
def getPerson(arg, PersonsList):
    suggestions =  Person.findSimiliarNames(arg,PersonsList)
    if suggestions == None:
        print("Nie znaleziono podanej osoby")
    elif len(suggestions) == 1:
        printPersonData(suggestions[0][1])
    else:
        for index, suggestion in suggestions:
            print(str(index) + ": " + suggestion.lastname+" "+suggestion.firstname)
        print("Wprowadź numer odpowiadający osobie")
        i = int(input())
        if i <= len(suggestions) and i > 0:
            printPersonData(suggestions[i-1][1])

def distance(a,b):
    n, m = len(a), len(b)
    if n > m:
        a,b = b,a
        n,m = m,n
    current = range(n+1)
    for i in range(1,m+1):
        previous, current = current, [i]+[0]*m
        for j in range(1,n+1):
            add, delete = previous[j]+1, current[j-1]+1
            change = previous[j-1]
            if a[j-1] != b[i-1]:
                change = change + 1
            current[j] = min(add, delete, change)

    return current[n]



def Main():
    data = Parser.ParseAll()
    input1 = input()
    print(input1)
    split_input =  input1.split(" ", 1)
    function = split_input[0]
    arg = split_input[1]
    while function.lower() != "stop":
        if function.lower() == "osoba":
            getPerson(arg, data[0])
        input1 = input()
        print(input1)
        split_input =  input1.split(" ", 1)
        function = split_input[0]
        arg = split_input[1]

    exRoom = Room.findByString("D17 4.30", data[1])
    #for rooms in RoomsList:
    #printPersonData(Person.findByName("Gajęcki", "Marek", data[0]))

if __name__=="__main__":
    Main()




