'Program to generate freePBX report for Comercial Hnos. Barreto
'1) Web scraper to get data. 
'2) Manipulate data
'==============================================================================================
'**GLOBAL VAR**
Dim HatilloQueues(9)
'Below line is the declaration of array variables to store the Hatillo Data.
Dim HatilloNoAnswer(10000)
Dim HatilloCompleted(10000)
Dim HatilloEnterQ(10000)
Dim HatilloAbandon(10000)

Dim AllNoAnswer(10000)
Dim AllCompleted(10000)
Dim AllEnterQ(10000)
Dim AllAbandon(10000)
'---------------------------------------
'==============================================================================================
'**WEB scrape**

Dim ie, frm, StoredRawData

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate "http://10.5.1.2/admin/config.php?display=logfiles"
ie.Visible=false 
WScript.Sleep 20000 'Wait 8 seconds

ie.document.getElementById("logfile").Value = "21" 'Chooses Que_Log value
'ie.document.getElementById("logfile").Value = "51" 'Chooses Que_Log value

ie.document.getElementById("lines").Value = "5000000000000"  'Select amount of lines

ie.document.getElementById("show").Click 'Executes filter

WScript.Sleep 8000 'Wait 8 seconds

StoredRawData = ie.document.getElementById("log_view").innerText 'Stores the desired data

ie.Quit 'Closes Internet Explorer

'MsgBox StoredRawData  'Para mosrtrar la data cruda

'DataLines = Split(StoredRawData,vbCrLf)

'lineNb = UBound(DataLines) 'Used to count the ammount of lines in the data extract.

'MsgBox lineNb 'Displays the ammount of lines gathered
'MsgBox DataLines(1)

'ManageData()
FileSave()
WScript.Quit

'====================================================================================
                    'Subs Section
'====================================================================================
Sub ManageData 'This section will manage the data pulled
 'First separate store calls by Queue number
    Dim Column2Check 'CallID Column

    For a = 0 to lineNb ' For to verify each line of data.
    Column2Check = Split(DataLines(a),"|")
  
    Next

End Sub
Sub FileSave 'Will be used to save the result on a txt file

' Create The Object
Set FSO = CreateObject("Scripting.FileSystemObject")

' How To Write To A File
Set File = FSO.CreateTextFile("Data.txt",True)
File.Write StoredRawData
File.Close

End Sub