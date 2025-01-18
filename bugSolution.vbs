Function GetObjectSafe(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    On Error Resume Next
    Set obj = CreateObject(progID)
    If Err.Number <> 0 Then
      Set obj = Nothing
    End If
  End If
  Set GetObjectSafe = obj
End Function

'This improved function includes more robust error handling.
'It checks for errors after both GetObject and CreateObject attempts.
'If both fail, it explicitly sets the object to Nothing.
Dim obj
Set obj = GetObjectSafe("Scripting.FileSystemObject")
If obj Is Nothing Then
  WScript.Echo "Could not get or create FileSystemObject"
Else
  WScript.Echo "FileSystemObject created successfully"
  'Do something with the object here
End If