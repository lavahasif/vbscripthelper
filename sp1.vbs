 If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
  end if
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

UpdateRegistry("256")