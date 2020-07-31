from bs4 import BeautifulSoup
from itertools import product
import pandas as pd

with open("C:/Users/Sencer Karakaya/Downloads/export.xls", "r", encoding='UTF-16') as excel:

    soup = BeautifulSoup(excel, "html.parser")
    info = soup.find("body").find_all("tr")[7:-5]

    g = []
    p = []
    a = []

    for q in info:

        course = q.find_all("td")[3:4]
        credit = q.find_all("td")[6:7]
        grade = q.find_all("td")[7:8]

        for y, t, u in product(course, credit, grade):

            try:
                g.append(y.text)
                p.append(int(t.text))
                a.append(u.text.strip().strip('"').strip("*"))
            finally:
                pass

    course = pd.Series(g)
    credit = pd.Series(p)
    grade = pd.Series(a)

    df = pd.concat([course, credit, grade], axis=1)
    df.columns = ["Courses", "Credits", "Grades"]
    df.index = df.index + 1

    for i in df[df["Credits"] == 0].index:
        df.drop(index=i, inplace=True)

    for j in df[(df["Grades"] == "--") | (df["Grades"] == "M")].index:
        df.drop(index=j, inplace=True)

    for dup in df[df.duplicated("Courses")].index:
        df.drop(index=dup, inplace=True)

    df.index = list(range(1, len(df.index) + 1))


def calculateGPA(df):

    total_credit = df["Credits"].sum()
    total_point = 0
    index = 0

    for grade in df["Grades"]:

        if grade == 'AA': grade = 4
        elif grade == 'BA': grade = 3.5
        elif grade == 'BB': grade = 3
        elif grade == 'CB': grade = 2.5
        elif grade == 'CC': grade = 2
        elif grade == 'DC': grade = 1.5
        elif grade == 'DD': grade = 1
        elif grade == 'FD': grade = 0.5
        else: grade = 0

        index += 1
        credit = df["Credits"][index]
        point = credit * grade
        total_point += point

        if index == len(df.index):
            break

    gpa = total_point / total_credit
    gpa = "{:.2f}".format(gpa)

    print(df)
    print(f'Total Credit: {total_credit}\nTotal Point: {total_point}\nGPA: {gpa}')


calculateGPA(df)


class Courses:

    def __init__(self,my_courses):
        self.my_courses = my_courses

    def add(self):
        name = input('Name: ')
        credit = int(input('Credit: '))
        grade = input('Grade: ')
        new_row = pd.Series([name,credit,grade],index=self.my_courses.columns)
        self.my_courses = self.my_courses.append(new_row,ignore_index=True)
        self.my_courses.index = list(range(1, len(self.my_courses.index) + 1))
        calculateGPA(self.my_courses)

    def remove(self):
        index_number = int(input('Index Number: '))
        self.my_courses = self.my_courses.drop(self.my_courses.index[index_number-1])
        self.my_courses.index = list(range(1, len(self.my_courses.index) + 1))
        calculateGPA(self.my_courses)

    def change(self):
        index_number = int(input('Index Number: '))
        new_grade = input('New Grade: ')
        self.my_courses.at[index_number,"Grades"] = new_grade
        calculateGPA(self.my_courses)


c = Courses(df)

while True:

    ctrl = input("\nPress 'E' To Continue!")
    enter = input("\n1- Add Course\n2- Remove Course\n3- Change Course Grade\n")

    if ctrl.lower() != "e":
        break
    elif enter == "1":
        c.add()
    elif enter == "2":
        c.remove()
    elif enter == "3":
        c.change()
    else:
        pass
