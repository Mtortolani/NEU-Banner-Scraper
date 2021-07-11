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
allClasses = ("Management",
              "Management Science",
              "Organizational Behavior", 
              "Suppy Chain Management", 
              "International Business", 
              "Strategy")

for classType in allClasses:
    for semester in semesters:
        emptyFile = '/Users/marcotortolani/Desktop/RA work for Bart/'+classType+' '+semester+'.json'
        #open(emptyFile,'a').close()