import bs4 as bs
import json
from openpyxl import Workbook
from openpyxl.styles import Font
 
 
print("start")



def BannerToExcel(subject: str, semesters: tuple):
    workbook = Workbook()

    for semester in semesters:
        with open(subject + " " + semester + ".json", "r") as file:
            # turn json file into big string
            content = file.read()
            # make replacements to content string
            clean = content.replace("&#39;", "\'")
            clean = clean.replace("&amp;", "AAA")
            #clean = clean.replace("amp;", "")
            # turn string into json formatted dictonary
            data = json.loads(clean)["data"]
        
        courseData = tuple(((data[i]["courseTitle"],
                            data[i]["subjectDescription"],
                            int(data[i]["courseNumber"]),
                            int(data[i]["sequenceNumber"]),
                            int(data[i]["creditHourLow"]),
                            semester,
                            int(data[i]["faculty"][0]["courseReferenceNumber"]),
                            data[i]["faculty"][0]["displayName"],
                            data[i]["campusDescription"],
                            max(0,int(data[i]["seatsAvailable"])),
                            int(data[i]["maximumEnrollment"]))  for i in range(0,len(data))))
        
        
        # Initialize Excel Workbook
        sheet = workbook.create_sheet(semester)
        sheet.title = semester
        
        
        
        sheet.append(("Title", 
                    "Subject",
                    "Course #",
                    "Section #",
                    "Hours",
                    "CRN",
                    "Term",
                    "Instructor",
                    "Campus", 
                    "Status - remaining seats",
                    "Status - total seats"))
        # Make first row bold
        for item in sheet[1]:
            item.font = Font(bold=True)
        
        # Add rows of data into the worksheet
        for row in courseData:
            # for item in row:
            #     if type(item) == str:
            #         item = item.replace("Finance", "YYYY")
            #         print(item)
            sheet.append(row)


    # Save Excel Sheet
    workbook.save(filename="RA Bart " + subject + ".xlsx")


def main():
    semesters = ("Fall 2015",
                 "Spring 2016",
                 "Fall 2016",
                 "Spring 2017",
                 "Fall 2017",
                 "Spring 2018",
                 "Fall 2018",
                 "Spring 2019",
                 "Fall 2019",
                 "Spring 2020",
                 "Fall 2020", 
                 "Spring 2021")
    BannerToExcel(subject="Finance & Insurance", semesters=semesters)

main()