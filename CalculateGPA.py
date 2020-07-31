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

    print(f'Total Credit: {total_credit}\nTotal Point: {total_point}\nGPA: {gpa}')
