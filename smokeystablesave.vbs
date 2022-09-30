
'wscript.echo DateString(now())
table=""



Function openFolderFile(path)
On Error Resume Next
set WSshell = createobject("wscript.shell")
WSshell.run path, 1
Set WSshell = nothing
On Error Goto 0
End Function
Function openFile(path)
On Error Resume Next
 set objShell = CreateObject("Shell.Application")

    objShell.ShellExecute path, "", "", "open", 1

    set objShell = nothing
On Error Goto 0
End Function

Function openSqlconfig()
openFolderFile("C:\\Windows\\SysWOW64\\SQLServerManager.msc")
openFolderFile("C:\\Windows\\SysWOW64\\SQLServerManager10.msc")
openFolderFile("C:\\Windows\\SysWOW64\\SQLServerManager15.msc")
openFolderFile("C:\\Windows\\SysWOW64\\SQLServerManager12.msc")

end Function
Function openSqlstudio()


openFile("C:\Program Files (x86)\Microsoft SQL Server\90\Tools\Binn\VSShell\Common7\IDE\ssmsee.exe")
openFile("C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe"  )
openFile("C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe"  )
openFile("C:\Program Files\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe"  )
openFile("C:\Program Files (x86)\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe")
openFile("C:\Program Files\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"  )
openFile("C:\Program Files (x86)\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"  )

end Function

function SAmpleExcel()
Dim Conn
Dim RS
Dim SQL
SQL = "SELECT PersonID, FirstName, LastName FROM [TestDB].[dbo].[Persons]"
Set Conn = CreateObject("ADODB.Connection")
Conn.Open = "Provider=SQLOLEDB; Data Source=compname\SQL; Initial Catalog=DB; UID=usera; Integrated Security=SSPI"

Set RS = Conn.Execute(SQL)

Set Sheet = ActiveSheet
Sheet.Activate

Dim R
R = 1
While RS.EOF = False
  Sheet.Cells(R, 1).Value = RS.Fields(0)
  Sheet.Cells(R, 2).Value = RS.Fields(1)
  Sheet.Cells(R, 3).Value = RS.Fields(2)
  RS.MoveNext
  R = R + 1
Wend

RS.Close
Conn.Close
end function
Function DateString(dDate)
    DateString = Year(dDate)& right("_" & Month(dDate),2) & right("_" & Day(dDate),2) & right("_" & Hour(dDate),2) & right("0" & Minute(dDate),2) & right("0" & second(dDate),2)
End Function

Function query(sql)
Const DB_CONNECT_STRING = "Provider=SQLOLEDB.1;Data Source=.\sqlexpress;Initial Catalog=vhotelhdd;user id ='sa';password='997755'"
Set myConn = CreateObject("ADODB.Connection")
Set myCommand = CreateObject("ADODB.Command" )
myConn.Open DB_CONNECT_STRING
Set myCommand.ActiveConnection = myConn
myCommand.CommandText = sql
myCommand.Execute
myConn.Close
end Function

Function backups()
sql="BACKUP DATABASE vhotelhdd TO DISK = 'D:\vhotelhdd" &DateString(now())& ".bak'"
query(sql)

end Function
Function MoveFromOldtoNew_Item()
sql="insert into item_registration_old SELECT TOP 1000 [Usercode] ,[ItemName] ,[Catagory] ,[acRate] ,[acdRate] ,[nonacRATE] ,[lessitem] ,[stockless] ,[prate] ,[PARCEL] ,[SUBCATAGORY] ,[gstper] ,[unit] ,[tax] ,[Slno123] ,[SLNO] ,[Image] ,[Active] FROM [vhotelhdd].[dbo].[Item_Registration]"
query (sql)

end Function

Function CreateItemExcel(tablename)

Set obj = createobject("ADODB.Connection") '�Creating an ADODB Connection Object
Set obj1 = createobject("ADODB.RecordSet") '�Creating an ADODB Recordset Object
strFileName = Replace(GetDownloadsPath(),"Desktop","Downloads\") & DateString(now())&".xls"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

' Set objWorkbook =objExcel.Workbooks.Open(strFileName)
' Set objWorkbook = objExcel.Workbooks.Open(strFileName)
 Set objWorkbook = objExcel.Workbooks.Add()
' set sheet=objWorkbook.WorkSheets("sheet1")

SQLStatement = "select * from " &tablename
' Note I moved your code around so that the sql statement 
' is defined before we open the recordset
ConnectionSqlServer="Provider=SQLOLEDB.1;Data Source=.\sqlexpress;Initial Catalog=vhotelhdd;user id ='sa';password='997755'"
obj.Open ConnectionSqlServer
obj1.Open SQLStatement, obj

For i = 0 To obj1.Fields.Count - 1
    
    with objWorkbook.WorkSheets("sheet1")
    .Cells(1, i + 1).Value = obj1.Fields(i).Name
    .Cells(1, i + 1).Font.Bold = True
    .Cells(1, i + 1).Interior.Pattern = xlSolid
   ' .Cells(1, i + 1).Interior.ThemeColor = xlThemeColorDark1
    .Cells(1, i + 1).Interior.ColorIndex = 3
    .Cells(1, i + 1).Font.Size = 16
    .Cells(1, i + 1).Font.Name = "Tahoma"


    end With
Next



objWorkbook.WorkSheets("sheet1").Range("A2").CopyFromRecordset obj1
obj1.Close

obj.close 

' objWorkbook.save()
objWorkbook.SaveAs(strFileName)
Set fso = CreateObject("Scripting.FileSystemObject")
Do While True
If fso.FileExists(strFileName)=True Then
 ' WScript.Echo strFileName & " exists."
'   objWorkbook.Close()
'   objExcel.Quit
  exit do
End If
loop

'If fso.FileExists(strFileName) Then
'  WScript.Echo filename & " exists."
'  objExcel.Quit
'End If

'WScript.Sleep 1000


end Function
Function GetDownloadsPath() 
   Dim sDesktopPath
Set objShell = Wscript.CreateObject("Wscript.Shell")
GetDownloadsPath = objShell.SpecialFolders("Desktop")
End Function
Function TruncateTable()
sql="truncate table item_registration"
query (sql)

end Function
Function IEButtons( )
    ' This function uses Internet Explorer to create a dialog.
    Dim objIE, sTitle, iErrorNum
    Dim defaultvalue 
    defaultvalue = 0

    ' Create an IE object
    Set objIE = CreateObject( "InternetExplorer.Application" )
    ' specify some of the IE window's settings
    objIE.Navigate "about:blank"
    sTitle="Click Any of the choice " & String( 80, "." ) 'Note: the String( 80,".") is to push "Internet Explorer" string off the window
    objIE.Document.title = sTitle
    objIE.MenuBar        = False
    objIE.ToolBar        = False
    objIE.AddressBar     = false
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 500
    objIE.Height         = 500
    ' Center the dialog window on the screen
    With objIE.Document.parentWindow.screen
        objIE.Left = (.availWidth  - objIE.Width ) \ 2
        objIE.Top  = (.availHeight - objIE.Height) \ 2
    End With
    ' Wait till IE is ready
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    

    ' Insert the HTML code to prompt for user input
    objIE.Document.body.innerHTML = "<div align=""center"">" & vbcrlf _
                                  & "<p><input type=""hidden"" id=""OK"" name=""OK"" value=""0"">" _
                                  & "<p><input type=""hidden"" id=""OK2"" name=""OK2"" value=""0"">" _
                                  & "<input type=""submit"" type=""hidden"" hidden name=""HI"" id=""HI""  value=""  Service Pack0   "" onClick=""VBScript:OK.value=0;VBScript:OK2.value=10""></p>" _
                                  & "<input type=""submit"" value=""  Backup "" onClick=""VBScript:OK.value=1""></p>" _
                                  & "<input type=""submit"" value=""   solve Item Problem   "" onClick=""VBScript:OK.value=2""></p>" _
                                  & "<input type=""submit"" value=""  Create Excel Sheet For Item "" onClick=""VBScript:OK.value=3""></p>" _
                                  & "<input type=""text"" id=""tab_name"" name=""tab_name"" value=""Item_Registration"" ></p>" _
                                   & "<input type=""submit"" value=""  Create Excel Sheet For Table "" onClick=""VBScript:OK.value=4""></p>" _
                                   & "<input type=""submit"" value=""  OpenDownloadFolder "" onClick=""VBScript:OK.value=5""></p>" _
                                   & "<input type=""submit"" value=""  Sqlconfig "" onClick=""VBScript:OK.value=6""></p>" _
                                   & "<input type=""submit"" value=""  SqlStudio "" onClick=""VBScript:OK.value=7""></p>" _
                                  & "<p><input type=""hidden"" id=""Cancel"" name=""Cancel"" value=""0"">" _
                                  & "<input type=""submit"" hidden id=""CancelButton"" value=""       Cancel       "" onClick=""VBScript:Cancel.value=-1""></p></div>"

    ' Hide the scrollbars
    objIE.Document.body.style.overflow = "auto"
    ' Make the window visible
    objIE.Visible = True
    ' Set focus on Cancel button
    objIE.Document.all.CancelButton.focus


    'CAVEAT: If user click red X to close IE window instead of click cancel, an error will occur.
    '        Error trapping Is Not doable For some reason
    On Error Resume Next
    dim s
    Do While objIE.Document.all.OK.value = 0 and objIE.Document.all.Cancel.value = 0
        WScript.Sleep 200
        iErrorNum=Err.Number
        table=objIE.Document.all.tab_name.value
    ' s= objIE.Document.all.OK2.value
        If iErrorNum <> 0   Then    'user clicked red X (or alt-F4) to close IE window
            IEButtons = 1000
            objIE.Quit
            Set objIE = Nothing
            Exit Function
        End if

        
    Loop
    On Error Goto 0

    objIE.Visible = False

    ' Read the user input from the dialog window
    IEButtons = objIE.Document.all.OK.value
    ' Close and release the object
    objIE.Quit
    Set objIE = Nothing
End Function

 If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  
  
  
  
  
  backup="1"
copyolditemtonew="2"
createExcelsheetitem="3"
tabletoExcel="4"
openDownload="5"
sqlconfig="6"
sqlstudio="7"

  selct=IEButtons()
if selct=backup Then
backups()
elseif selct=copyolditemtonew Then
CreateItemExcel("Item_Registration")
backup()
TruncateTable()
MoveFromOldtoNew_Item()

elseif selct=createExcelsheetitem Then
CreateItemExcel("Item_Registration")
elseif selct=opendownload Then

openFolderFile(Replace(GetDownloadsPath(),"Desktop","Downloads"))
elseif selct=sqlstudio Then
openSqlstudio()
elseif selct=sqlconfig Then
openSqlconfig()
elseif selct=tabletoExcel Then

CreateItemExcel(table)
end if

  
  
  
  WScript.Sleep 200
  WScript.Quit

  

End If



