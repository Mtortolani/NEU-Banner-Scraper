from typing import DefaultDict
import bs4 as bs
import json
from openpyxl import Workbook
from openpyxl.styles import Font
from collections import defaultdict
 
class BannerRunner:
    def __init__(self, subjects:tuple):
        #self.subjects = tuple()
        self.subjects = subjects
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
        # courseDataDict[subject][semester]
        self.courseDataDict = defaultdict(dict)
        self.workbooks = list()

    def TextCleaner(self, dirty:str):
        dirty = dirty.replace("&#39;", "\'")
        dirty = dirty.replace("&amp;", "AAA")
        return dirty



    def JsonToCourseData(self, semester:str, subject:str):
        with open(semester+'.json','r') as file:
            # turn json file into big string
            content = file.read()
            # make replacements to content string
            clean = self.TextCleaner(content)
            #clean = clean.replace("amp;", "")
            # turn string into json formatted dictonary
            data = json.loads(clean)["data"]
            courseData = list()
            for i in range(0,len(data)):
                if data[i]['subjectDescription']==subject:
                    courseData.append((data[i]["courseTitle"],
                                data[i]["subjectDescription"],
                                int(data[i]["courseNumber"]),
                                int(data[i]["sequenceNumber"]),
                                int(data[i]["creditHourLow"]),
                                data[i]["termDesc"].replace(" Semester", ""),
                                int(data[i]["faculty"][0]["courseReferenceNumber"]),
                                data[i]["faculty"][0]["displayName"],
                                data[i]["campusDescription"],
                                int(data[i]["seatsAvailable"]),
                                int(data[i]["maximumEnrollment"])))
                                      
            self.courseDataDict[subject][semester] = courseData
            
        
    
    def MoveAllJsonToCourseData(self):
        for semester in self.semesters:
            for subject in self.subjects:
                self.JsonToCourseData(semester, subject)
        
    def WidthSetter(worksheet):
        '''
        Citation: Bufke & Eli, Stack Overflow
        https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size
        '''
        dims = {}
        for row in worksheet.rows:
            for cell in row:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value)), 15))
                    
        for col, value in dims.items():
            worksheet.column_dimensions[col].width = value
        

    def MoveAllCourseDataToExcel(self):
        for subject in self.subjects:
            workbook = Workbook()
            for semester in self.semesters:
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
                for row in self.courseDataDict[subject][semester]:
                    sheet.append(row)



                BannerRunner.WidthSetter(sheet)
            


            # Save Excel File
            workbook.save(filename="RA Bart " + subject + ".xlsx")



def main():
    allClasses = ("Management", 
                  "Management Science", 
                  "Organizational Behavior", 
                  "Supply Chain Management", 
                  "International Business", 
                  "Strategy")
    Runner = BannerRunner(allClasses)
    Runner.MoveAllJsonToCourseData()
    Runner.MoveAllCourseDataToExcel()


    

main()