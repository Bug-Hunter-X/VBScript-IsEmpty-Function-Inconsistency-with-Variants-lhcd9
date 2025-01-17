Function MyFunc(param1)
  If VarType(param1) = vbEmpty Then
    param1 = ""
  ElseIf IsEmpty(param1) Then
    param1 = ""
  End If
  ' ... rest of the function
End Function