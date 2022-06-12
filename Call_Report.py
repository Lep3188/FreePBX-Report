# Developed by Luis Peralta at CHB
# Date created:  ‎April ‎28, ‎2021, ‏‎3:51:45 PM
# Completion Date: 

#Program search and opens the Data file, Sorts the data and create the missed calls report. 
from typing import Counter
import numpy as np 
import os
import os.path
import time
import smtplib
import datetime
from datetime import datetime

#========================================================
#      Functions
#========================================================

#Function to Run VBS file
def VBSRun():
   os.system("Call_Report.vbs")

def Main_Code():
#This is the function that manages the data
   RawDataGet = open(file_path, "r")

   DataSplit = RawDataGet.read().split('\n') #Stores each line into array

   DataAmmount = len(DataSplit) 
   RawDataGet.close()

   HatilloArray = np.array([]) #Declaration of unlimited array variable of Hatillo store
   LaresArray = np.array([]) #Declaration of unlimited array variable of Lares store
   AreciboArray = np.array([]) #Declaration of unlimited array variable of Arecibo store
   BloqueraArray = np.array([]) #Declaration of unlimited array variable of Arecibo store
   FarmaciaOne = np.array([]) #Declaration of unlimited array variable of Arecibo store
   FarmaciaTwo = np.array([]) #Declaration of unlimited array variable of Arecibo store
   trashData = np.array([])

   for x in DataSplit:
      WorkLine = x
      ColumnSearchSplit = WorkLine.split('|')

      if ColumnSearchSplit[2] == 'NONE':
         trashData = np.append(trashData,x)
      elif int(ColumnSearchSplit[2]) <= 608:
         HatilloArray = np.append(HatilloArray,x) #Adding to Array example
      elif(int(ColumnSearchSplit[2]) == 609):
         LaresArray = np.append(LaresArray,x) #Adding to Array example
      elif(610 <= int(ColumnSearchSplit[2]) <= 614):
         AreciboArray = np.append(AreciboArray,x) #Adding to Array example
      elif(int(ColumnSearchSplit[2]) == 615):
         BloqueraArray = np.append(BloqueraArray,x) #Adding to Array example
      elif(627 <= int(ColumnSearchSplit[2]) <= 628):
         FarmaciaOne = np.append(FarmaciaOne,x) #Adding to Array example
      elif(617 <= int(ColumnSearchSplit[2]) <= 626):
         FarmaciaTwo = np.append(FarmaciaTwo,x) #Adding to Array example

   #==========================================================
   #This section will sort the queue data by store
   #==========================================================

   HatilloEnterQueueArray = np.array([])
   HatilloRingNoAnswerArray = np.array([])
   HatilloConnectArray = np.array([])
   HatilloAbandonArray = np.array([])

   AreciboEnterQueueArray = np.array([])
   AreciboRingNoAnswerArray = np.array([])
   AreciboConnectArray = np.array([])
   AreciboAbandonArray = np.array([])

   LaresEnterQueueArray = np.array([])
   LaresRingNoAnswerArray = np.array([])
   LaresConnectArray = np.array([])
   LaresAbandonArray = np.array([])

   for x in HatilloArray:
      HatilloWorkLine = x

      HatilloDataSplit = HatilloWorkLine.split('|')
      if (HatilloDataSplit[4]) == "ENTERQUEUE":
         HatilloEnterQueueArray = np.append(HatilloEnterQueueArray,x) #Adding to Array example
      elif (HatilloDataSplit[4]) == "RINGNOANSWER":
         HatilloRingNoAnswerArray = np.append(HatilloRingNoAnswerArray,x) #Adding to Array example
      elif (HatilloDataSplit[4]) == "CONNECT":
         HatilloConnectArray = np.append(HatilloConnectArray,x) #Adding to Array example
      elif (HatilloDataSplit[4]) == "ABANDON":
         HatilloAbandonArray = np.append(HatilloAbandonArray,x) #Adding to Array example

   for x in LaresArray:
      LaresWorkLine = x

      LaresDataSplit = LaresWorkLine.split('|')
      if (LaresDataSplit[4]) == "ENTERQUEUE":
         LaresEnterQueueArray = np.append(LaresEnterQueueArray,x) #Adding to Array example
      elif (LaresDataSplit[4]) == "RINGNOANSWER":
         LaresRingNoAnswerArray = np.append(LaresRingNoAnswerArray,x) #Adding to Array example
      elif (LaresDataSplit[4]) == "CONNECT":
         LaresConnectArray = np.append(LaresConnectArray,x) #Adding to Array example
      elif (LaresDataSplit[4]) == "ABANDON":
         LaresAbandonArray = np.append(LaresAbandonArray,x) #Adding to Array example

   for x in AreciboArray:
      AreciboWorkLine = x

      AreciboDataSplit = AreciboWorkLine.split('|')
      if (AreciboDataSplit[4]) == "ENTERQUEUE":
         AreciboEnterQueueArray = np.append(AreciboEnterQueueArray,x) #Adding to Array example
      elif (AreciboDataSplit[4]) == "RINGNOANSWER":
         AreciboRingNoAnswerArray = np.append(AreciboRingNoAnswerArray,x) #Adding to Array example
      elif (AreciboDataSplit[4]) == "CONNECT":
         AreciboConnectArray = np.append(AreciboConnectArray,x) #Adding to Array example
      elif (AreciboDataSplit[4]) == "ABANDON":
         AreciboAbandonArray = np.append(AreciboAbandonArray,x) #Adding to Array example

   #==========================================================
   #End of QueueSort by store
   #==========================================================

   #Beggining of Missed call duplicate removal
   #===========================================

   #This cycle will separate de data and store only the CallID of missed calls

   HatilloCleanMissed = np.array([]) 
   HatilloCallIDMissed = np.array([]) 

   AreciboCleanMissed = np.array([]) 
   AreciboCallIDMissed = np.array([]) 

   LaresCleanMissed = np.array([]) 
   LaresCallIDMissed = np.array([]) 

   #*********************************************************************
   #Hatillo
   for x in HatilloRingNoAnswerArray:
      currentData = x.split('|')
      if currentData[1] not in HatilloCallIDMissed:
         HatilloCallIDMissed = np.append(currentData[1], HatilloCallIDMissed)
         HatilloCleanMissed = np.append(HatilloCleanMissed, x)
   #**********************************************************************
   #Lares
   for x in LaresRingNoAnswerArray:
      currentData = x.split('|')
      if currentData[1] not in LaresCallIDMissed:
         LaresCallIDMissed = np.append(currentData[1], LaresCallIDMissed)
         LaresCleanMissed = np.append(LaresCleanMissed, x)
   #**********************************************************************
   #Arecibo
   for x in AreciboRingNoAnswerArray:
      currentData = x.split('|')
      if currentData[1] not in AreciboCallIDMissed:
         AreciboCallIDMissed = np.append(currentData[1], AreciboCallIDMissed)
         AreciboCleanMissed = np.append(AreciboCleanMissed, x)

   #----------------------------------------------------------------------
   #Area to search for real Missed calls, this is by comparing CleanMissedCAlls to the Connected ones by their ID. 
   HatilloNOtAnswered = np.array([]) 
   HatilloConnectID = np.array([]) 

   for x in HatilloConnectArray:
      SplitedDataConnect = x.split("|")
      HatilloConnectID = np.append(HatilloConnectID, SplitedDataConnect[1])

   for x in HatilloCleanMissed:
      SplitedDataMissed = x.split("|")

      if SplitedDataMissed[1] not in HatilloConnectID:
         HatilloNOtAnswered = np.append(HatilloNOtAnswered,x)

   #==========================================================
   AreciboNOtAnswered = np.array([]) 
   AreciboConnectID = np.array([]) 

   for x in AreciboConnectArray:
      SplitedDataConnect = x.split("|")
      AreciboConnectID = np.append(AreciboConnectID, SplitedDataConnect[1])

   for x in AreciboCleanMissed:
      SplitedDataMissed = x.split("|")

      if SplitedDataMissed[1] not in AreciboConnectID:
         AreciboNOtAnswered = np.append(AreciboNOtAnswered,x)

   #==========================================================
   LaresNOtAnswered = np.array([]) 
   LaresConnectID = np.array([]) 

   for x in LaresConnectArray:
      SplitedDataConnect = x.split("|")
      LaresConnectID = np.append(LaresConnectID, SplitedDataConnect[1])

   for x in LaresCleanMissed:
      SplitedDataMissed = x.split("|")

      if SplitedDataMissed[1] not in LaresConnectID:
         LaresNOtAnswered = np.append(LaresNOtAnswered,x)

   #==========================================================
   #Attempt to create the final report.
   #First we go by each Not answered call, split it and search for phone number inside enterqueue data. Then we-
   # make a string with all the values and display it like a table. Timestamp gets converted to Date.

   HatilloReportData = np.array([])
   HatilloAbandonReportData = np.array([])

   for x in HatilloNOtAnswered: 
      x2 = x.split("|")
      for j in HatilloEnterQueueArray:
         j2 = j.split("|")

         if x2[1] == j2[1]:
            dt_object = datetime.fromtimestamp(int(x2[0]))
            result = (str(x2[1])+ "||"+ str(j2[6])+ "||"+  str(dt_object))

            HatilloReportData = np.append(HatilloReportData, result)
   for x in HatilloAbandonArray: 
      x2 = x.split("|")
      for j in HatilloEnterQueueArray:
         j2 = j.split("|")

         if x2[1] == j2[1]:
            dt_object = datetime.fromtimestamp(int(x2[0]))
            resultAbandon = (str(x2[1])+ "|| "+ str(j2[6])+ " ||"+  str(dt_object))

            HatilloAbandonReportData = np.append(HatilloAbandonReportData, resultAbandon)
   #Arecibo
   AreciboReportData = np.array([])
   AreciboAbandonReportData = np.array([])
   for x in AreciboNOtAnswered: 
      x2 = x.split("|")
      for j in AreciboEnterQueueArray:
         j2 = j.split("|")

         if x2[1] == j2[1]:
            dt_object = datetime.fromtimestamp(int(x2[0]))
            result = (str(x2[1])+ "||"+ str(j2[6])+ "||"+  str(dt_object))

            AreciboReportData = np.append(AreciboReportData, result)
   for x in AreciboAbandonArray: 
      x2 = x.split("|")
      for j in HatilloEnterQueueArray:
         j2 = j.split("|")

         if x2[1] == j2[1]:
            dt_object = datetime.fromtimestamp(int(x2[0]))
            resultAbandon = (str(x2[1])+ "|| "+ str(j2[6])+ " ||"+  str(dt_object))

            AreciboAbandonReportData = np.append(AreciboAbandonReportData, resultAbandon)
   #==========================================================
   #Abandon Calls Cleanup based in calls that where missed
   #==========================================================

   RealAbandonsHatillo = np.array([])
   AbandonPartialStored = np.array([])
   HatilloMissedArrayIDHold = np.array([])

   for x in HatilloReportData:
      x2 = x.split("||")
      HatilloMissedArrayIDHold =np.append(HatilloMissedArrayIDHold, x2[0])
   for x in HatilloAbandonReportData:
      x2 = x.split("||")
      #print (x2[0])
      if x2[0] not in HatilloMissedArrayIDHold:
         RealAbandonsHatillo = np.append(RealAbandonsHatillo, x)
   #==========================================================
   #Arecibo
   RealAbandonsArecibo = np.array([])
   AbandonPartialStored = np.array([])
   AreciboMissedArrayIDHold = np.array([])

   for x in AreciboReportData:
      x2 = x.split("||")
      AreciboMissedArrayIDHold =np.append(AreciboMissedArrayIDHold, x2[0])
   for x in HatilloAbandonReportData:
      x2 = x.split("||")
      #print (x2[0])
      if x2[0] not in AreciboMissedArrayIDHold:
         RealAbandonsArecibo = np.append(RealAbandonsArecibo, x)
    #==========================================================     
   Totales= open("Totales.txt","w")

   Totales.write("Trash Data: "+ str(len(trashData)) + "\n")
   Totales.write("Total Records: "+ str((DataAmmount)) + "\n")
   Totales.write("======================================="+ "\n")
   Totales.write("Records Obtenidos por tienda"+ "\n")
   Totales.write("Records de Hatillo: "+  str(len(HatilloArray))+ "\n")
   Totales.write("Records de Lares: "+  str(len(LaresArray))+ "\n")
   Totales.write("Records de Arecibo: "+  str(len(AreciboArray))+ "\n")
   Totales.write("Records de Bloquera: "+  str(len(BloqueraArray))+ "\n")
   Totales.write("Records de Farmacia 1: "+  str(len(FarmaciaOne))+ "\n")
   Totales.write("Records de Farmacia 2: "+  str(len(FarmaciaTwo))+ "\n")
   Totales.write("======================================="+ "\n")
   #==========================================================

   Totales.write("Records de Hatillo:\n")
   Totales.write("Connected Calls: "+ str(len(HatilloConnectArray))+ "\n")
   Totales.write("EnterQueue Calls: "+ str(len(HatilloEnterQueueArray))+ "\n")
   Totales.write("Missed Calls RAW: "+ str(len(HatilloRingNoAnswerArray))+ "\n")
   Totales.write("Real Missed Calls: "+ str(len(HatilloNOtAnswered))+ "\n")
   Totales.write("Abandon Calls: "+ str(len(RealAbandonsHatillo))+ "\n")

   Totales.write("======================================="+ "\n")

   Totales.write("Records de Lares:\n")
   Totales.write("Connected Calls: "+  str(len(LaresConnectArray))+ "\n")
   Totales.write("EnterQueue Calls: "+  str(len(LaresEnterQueueArray))+ "\n")
   Totales.write("Missed Calls RAW: "+  str(len(LaresRingNoAnswerArray))+ "\n")
   Totales.write("Missed Calls Clean: "+  str(len(LaresNOtAnswered))+ "\n")
   Totales.write("Abandon Calls: "+  str(len(LaresAbandonArray))+ "\n")
   Totales.write("======================================="+ "\n")

   Totales.write("Records de Arecibo:"+ "\n")
   Totales.write("Connected Calls: "+  str(len(AreciboConnectArray))+ "\n")
   Totales.write("EnterQueue Calls: "+  str(len(AreciboEnterQueueArray))+ "\n")
   Totales.write("Missed Calls RAW: "+  str(len(AreciboRingNoAnswerArray))+ "\n")
   Totales.write("Missed Calls Clean: "+  str(len(AreciboNOtAnswered))+ "\n")
   Totales.write("Abandon Calls: "+  str(len(RealAbandonsArecibo))+ "\n")
   Totales.write("======================================="+ "\n")
   Totales.close()

   #=====================================================
   #Create txt files to store abandon, Missed and General Data
   Perdidas= open("Perdidas.txt","w")
   Perdidas.write("Reporte de llamadas perdidas: " + str(datetime.now()))
   Perdidas.write("\n")
   Perdidas.write("===============================================" + "\n")
   for i in HatilloReportData:
      Perdidas.write(i + "\n")
   Perdidas.close() 
   #=======================================================
   #Create txt files to store abandon, Missed and General Data
   Abandono= open("Abandonadas.txt","w")
   Abandono.write("Reporte de llamadas Abandonadas Hatillo: " + str(datetime.now()))
   Abandono.write("\n")
   Abandono.write("===============================================" + "\n")
   for i in RealAbandonsHatillo:
      Abandono.write(i + "\n")
   Abandono.close() 
   #=======================================================
   #Send email to receptionist of Hatillo

   from email.mime.multipart import MIMEMultipart
   from email.mime.text import MIMEText
   mail_content = "Llamadas Perdidas: \n===============================================\n"+ "\n".join(HatilloReportData) + "\n" + "\nLlamadas Abandonadas: \n===============================================\n"+"\n".join(RealAbandonsHatillo)
   #The mail addresses and password
   sender_address = 'admin@comercialbarreto.com'
   sender_pass = 'jjejwvjnlttopxtk'
   receiver_address = 'empleados@comercialbarreto.com'

   #Setup the MIME
   message = MIMEMultipart()
   message['From'] = sender_address
   message['To'] = receiver_address
   message['Subject'] = 'Reporte de llamadas Hatillo ' +  str(datetime.now())  #The subject line
   message.attach(MIMEText(mail_content, 'plain')) #The body and the attachments for the mail

   #Create SMTP session for sending the mail
   session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
   session.starttls() #enable security
   session.login(sender_address, sender_pass) #login with mail_id and password
   text = message.as_string()
   session.sendmail(sender_address, receiver_address, text)
   session.quit()
   print('Hatillo Mail Sent')

#=======================================================
   #Send email to Manager of Arecibo

   from email.mime.multipart import MIMEMultipart
   from email.mime.text import MIMEText
   mail_content = "Llamadas Perdidas: \n===============================================\n"+ "\n".join(AreciboReportData) + "\n" + "\nLlamadas Abandonadas: \n===============================================\n"+"\n".join(RealAbandonsArecibo)
   #The mail addresses and password
   sender_address = 'admin@comercialbarreto.com'
   sender_pass = 'jjejwvjnlttopxtk'
   receiver_address = 'empleados@comercialbarreto.com'

   #Setup the MIME
   message = MIMEMultipart()
   message['From'] = sender_address
   message['To'] = receiver_address
   message['Subject'] = 'Reporte de llamadas Arecibo ' +  str(datetime.now())  #The subject line
   message.attach(MIMEText(mail_content, 'plain')) #The body and the attachments for the mail

   #Create SMTP session for sending the mail
   session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
   session.starttls() #enable security
   session.login(sender_address, sender_pass) #login with mail_id and password
   text = message.as_string()
   session.sendmail(sender_address, receiver_address, text)
   session.quit()
   print('Arecibo Mail Sent')
   #========================================================
   #      End of program
   #========================================================
   os.remove("Data.txt") #Deletes the Extracted data file

#=========================================
#     Start of functions execution 
#=========================================

VBSRun() #Executes the function that will run a vb script to get the data (Web Scrape)

#=========================================
#      Wait for file to exist       
#=========================================

while True:
   file_path ="Data.txt"
   FileCheckCounter = 0
   while not os.path.exists(file_path):
      time.sleep(5) 
      FileCheckCounter = FileCheckCounter + 1
      
      if FileCheckCounter == 3: #If 15 seconds pass and file does not exist, exectes the funcion again.
         VBSRun()
   FileCheckCounter = 0
   #=========================================
   Main_Code() #Executes the function that will manage the data
   time.sleep(1800)