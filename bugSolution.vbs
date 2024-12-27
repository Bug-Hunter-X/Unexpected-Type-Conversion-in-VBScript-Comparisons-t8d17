Function f(a, b)
  'Explicitly convert both inputs to numbers for consistent comparison
  Dim aNum, bNum
  aNum = CDbl(a)
  bNum = CDbl(b)

  If aNum > bNum Then
    MsgBox "a is greater than b"
  ElseIf aNum < bNum Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
end function

Call f("5", 10)  ' Test with string and number
Call f(5, "10")  ' Test with number and string
Call f(5, 5)     ' Test with numbers
Call f("5", "5") 'Test with strings