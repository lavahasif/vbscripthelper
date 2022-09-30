


Function IEButtons( )
    ' This function uses Internet Explorer to create a dialog.
    Dim objIE, sTitle, iErrorNum
    Dim defaultvalue 
    defaultvalue = 0

    ' Create an IE object
    Set objIE = CreateObject( "InternetExplorer.Application" )
    ' specify some of the IE window's settings
    objIE.Navigate "about:blank"
    sTitle="Make your choice " & String( 80, "." ) 'Note: the String( 80,".") is to push "Internet Explorer" string off the window
    objIE.Document.title = sTitle
    objIE.MenuBar        = False
    objIE.ToolBar        = False
    objIE.AddressBar     = false
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 250
    objIE.Height         = 280
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
                                  & "<input type=""submit"" name=""HI"" id=""HI""  value=""  Service Pack0   "" onClick=""VBScript:OK.value=0;VBScript:OK2.value=10""></p>" _
                                  & "<input type=""submit"" value=""  Service Pack1 "" onClick=""VBScript:OK.value=256""></p>" _
                                  & "<input type=""submit"" value=""   Service Pack2   "" onClick=""VBScript:OK.value=512""></p>" _
                                  & "<input type=""submit"" value=""   Service Pack3   "" onClick=""VBScript:OK.value=768""></p>" _
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
     s= objIE.Document.all.OK2.value
        If iErrorNum <> 0   Then    'user clicked red X (or alt-F4) to close IE window
            IEButtons = 1000
            objIE.Quit
            Set objIE = Nothing
            Exit Function
        End if

          If s = 10   Then    'user clicked red X (or alt-F4) to close IE window
            IEButtons = 0
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
Function INewValue(value)
if value=256 then
     INewValue="sp1"

     elseif value=512 then
     INewValue="sp2"
          elseif value=768 then
     INewValue="sp3"
     else
       INewValue="sp0"
     end if
end Function

Function UpdateRegistry(new_value)
Dim sKey, bFound,sValue_new,sValue
skey = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Windows\CSDVersion"

with CreateObject("WScript.Shell")
 on error resume next            ' turn off error trapping
   sValue = .regread(sKey)  
   .RegWrite skey,new_value,  "REG_DWORD"
    ' read attempt
    sValue_new = .regread(sKey)  
   bFound = (err.number = 0)     ' test for success
 on error goto 0                 ' restore error trapping
end with

If bFound Then
  sValue_new=INewvalue(sValue_new)
  sValue=INewvalue(sValue)
   
 MsgBox  "Registry Updated Version:" & sValue &  "Changed to:" & sValue_new
Else
 MsgBox  "Nope, it doesn't exist. You Selected:"&new_value  
End If
End Function


 If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit

  

End If
  sp=IEButtons( )
if sp <> 1000 then
   UpdateRegistry(sp)
 end if