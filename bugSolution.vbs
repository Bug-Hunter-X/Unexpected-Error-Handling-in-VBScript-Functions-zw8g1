Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbError, , "Missing Parameters"
    Exit Function 'Important: exit the function after raising an error.
  End If
  ' ... rest of the function ...
End Function

' Example usage showing improved error handling.
On Error GoTo ErrorHandler

Dim result: result = MyFunction()
MsgBox "Result: " & result

Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description
  Err.Clear
End Sub