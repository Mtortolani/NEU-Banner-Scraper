import bs4 as bs
import json
from openpyxl import Workbook
from openpyxl.styles import Font
 
class BannerRunner:
    def __init__(self, subject):
        #self.subjects = tuple()
        self.subject = subject
        self.semesters =("Fall 2015",
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
        self.courseDataDict = {}
        self.workbook = Workbook()
    
    def JsonToCourseDataBySemester(self, semester:str)->tuple:
        with open(self.subject+ " " + semester + ".json", "r") as file:
            # turn json file into big string
            content = file.read()
            # make replacements to content string
            clean = content.replace("&#39;", "\'")
            clean = clean.replace("&amp;", "AAAAAAAAAA1")
            #clean = clean.replace("amp;", "")
            # turn string into json formatted dictonary
            data = json.loads(clean)["data"]
            
            courseData = tuple(((data[i]["courseTitle"],
                        data[i]["subjectDescription"],
                        int(data[i]["courseNumber"]),
                        int(data[i]["sequenceNumber"]),
                        int(data[i]["creditHourLow"]),
                        data[i]["termDesc"].replace(" semester", ""),
                        int(data[i]["faculty"][0]["courseReferenceNumber"]),
                        data[i]["faculty"][0]["displayName"],
                        data[i]["campusDescription"],
                        int(data[i]["seatsAvailable"]),
                        int(data[i]["maximumEnrollment"]))  for i in range(0,len(data))))
            self.courseDataDict[semester] = courseData
    
    def MoveAllJsonToCourseData(self):
        for semester in self.semesters:
            self.JsonToCourseData(semester)
        
    
    def MoveAllCourseDataToExcel(self):
        for semester in self.semesters:
            sheet = self.workbook.create_sheet(semester)
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
            for row in self.courseDataDict[semester]:
                sheet.append(row)


        # Save Excel Sheet
        self.workbook.save(filename="RA Bart " + self.subject + ".xlsx")



def main():
    allClasses = ("Management", 
                  "Management Science", 
                  "Organizational Behavior", 
                  "Supply Chain Management", 
                  "International Business", 
                  "Strategy")
    for oneClass in allClasses:
        Finance_and_Insurance = BannerRunner(oneClass)
        Finance_and_Insurance.MoveAllJsonToCourseData()
        Finance_and_Insurance.MoveAllCourseDataToExcel()
    

main()