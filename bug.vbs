Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Set obj = CreateObject(progID)
  End If
  Set GetObject = obj
End Function

'This function demonstrates an uncommon error.  
'It attempts to use GetObject to get an object. If it fails it will create the object
'The error is that if CreateObject fails it will still return an object, but an empty one
'This will lead to unexpected behavior in your program
Dim obj
Set obj = GetObject("Scripting.FileSystemObject")
If obj Is Nothing Then
  WScript.Echo "Could not get or create FileSystemObject"
Else
  WScript.Echo "FileSystemObject created successfully"
  'Do something with the object here
End If