from JournalInfo import JournalInfo
import pymysql
from pymysql.err import MySQLError
from array import array
from audioop import reverse
import operator
import xlwt
from tempfile import TemporaryFile
 
class JournalAlgorithm:
     
    def computeFiveYrAvg(category,year,noYears):
        arrayOfTitles=[]
        fiveYrAvg=[]
        sortedList=[]
        try:
            lastYear = year -5
            blwYear = year+1
            db = pymysql.connect(host='10.36.0.145',    
                     user='isajournal406',         
                     password='skip406', 
                     db='isajournals',
                     cursorclass=pymysql.cursors.DictCursor
                     )
            with db.cursor() as cursor:
                query = "SELECT DISTINCT `Full JOURNAL TITLE` FROM journal WHERE category = '" + category + "';"
                cursor.execute(query)
                result=cursor.fetchall()
                data=result
         
            for j in range (0,len(data)):
                arrayOfTitles.append(data[j]['Full JOURNAL TITLE'])
         
                journal = []
         
            for i in range (0,arrayOfTitles.__len__()):
                jTitle = arrayOfTitles[i]
                journalImp=[]
                query2 = "SELECT `5-year impact factor` FROM journal WHERE category = '" + category + "' AND "+ "`FULL JOURNAL TITLE` = '" + jTitle + "' AND Year > " + str(lastYear) + " AND Year < " + str(blwYear) + ";"
                with db.cursor() as cursor:
                    cursor.execute(query2)
                    result2=cursor.fetchall()
                    data2=result2
                    impactInd = 0
 
            
                for impactInd in range(0,5):
                    if(impactInd>=len(data2)):
                        journalImp.append(0)
                    else:
                        if data2[impactInd]["5-year impact factor"] == "Not Available":
                            journalImp.append(0)
                        else:
                            journalImp.append(data2[impactInd]["5-year impact factor"])
                total=0.0
                numNonZero=0.0
                if float(journalImp[0]) > 0:
                    total=total+float(journalImp[0])
                    numNonZero=numNonZero+1
                if float(journalImp[1]) > 0:
                    total=total+float(journalImp[1])
                    numNonZero=numNonZero+1
                if float(journalImp[2]) > 0:
                    total=total+float(journalImp[2])
                    numNonZero=numNonZero+1
                if float(journalImp[3]) > 0:
                    total=total+float(journalImp[3])
                    numNonZero=numNonZero+1
                if float(journalImp[4]) > 0:
                    total=total+float(journalImp[4])
                    numNonZero=numNonZero+1            
                if numNonZero>=noYears:
                    ave = total/numNonZero
                    fiveYrAvg.append(ave)
                else:
                    fiveYrAvg.append(0)
                

                
                rank = 0 
                groupID=''
                   
                journal.append(JournalInfo(rank,groupID,jTitle, category, year, float(journalImp[0]), float(journalImp[1]), float(journalImp[2])
                        , float(journalImp[3]), float(journalImp[4]), float(fiveYrAvg[i])))
                
        except MySQLError as e:
            print (e)
        db.close()
        journal.sort()
        numZeroRows = 0
        
        for g in range(0,len(journal)):
            ave = getattr(journal[g], 'fiveYrAvg')
            if ave==0:
                numZeroRows=numZeroRows+1
        actualLength = len(journal)-numZeroRows
            
        quint = actualLength/5
        median = actualLength/2
        group1Cutoff=getattr(journal[int(quint)], 'fiveYrAvg')
        group2Cutoff=getattr(journal[int(median)], 'fiveYrAvg')
        
        rank = 0   
        for g in range(0,len(journal)):
            groupID = 0
            average = getattr(journal[g], 'fiveYrAvg')
            if average >=group1Cutoff:
                groupID = 'I'
            elif average >=group2Cutoff: 
                groupID = 'II'
            else:
                groupID = 'III'
            rank = rank+1
            setattr(journal[g], 'rank', rank)
            setattr(journal[g], 'groupID', groupID)
        for e in journal:
            print(e,'\n')
        book = xlwt.Workbook()
        sheet1 = book.add_sheet('sheet1')
        
        sheet1.write(0,0,'Grouping')
        sheet1.write(0,1,'Rank')
        sheet1.write(0,2,'Title')
        sheet1.write(0,3,year-4)
        sheet1.write(0,4,year-3)
        sheet1.write(0,5,year-2)
        sheet1.write(0,6,year-1)
        sheet1.write(0,7,year)
        sheet1.write(0,8,"fiveYrAvg")
        
        for i in range(0,len(journal)):
            r=getattr(journal[i], 'rank')
            g=getattr(journal[i], 'groupID')
            t=getattr(journal[i], 'title')
            JPM0=getattr(journal[i], 'yr1JIF')
            JPM1=getattr(journal[i], 'yr2JIF')
            JPM2=getattr(journal[i], 'yr3JIF')
            JPM3=getattr(journal[i], 'yr4JIF')
            JPM4=getattr(journal[i], 'yr5JIF')
            a=getattr(journal[i],'fiveYrAvg')
            
            
            sheet1.write(i+1,0,g)
            sheet1.write(i+1,1,r)
            sheet1.write(i+1,2,t)
            sheet1.write(i+1,3,JPM0)
            sheet1.write(i+1,4,JPM1)
            sheet1.write(i+1,5,JPM2)
            sheet1.write(i+1,6,JPM3)
            sheet1.write(i+1,7,JPM4)
            sheet1.write(i+1,8,a)

        name = "random.xls"
        book.save(name)
        book.save(TemporaryFile())

     
    computeFiveYrAvg("Management", 2013,4)
#     print (a) 