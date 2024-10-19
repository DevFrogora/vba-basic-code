Dim Global_Return
MsgBox "" & MessageBox_Demo()

Function MessageBox_Demo() As Integer
   'Message Box with just prompt message
   'MsgBox ("Welcome")
   
   'Message Box with title, yes no and cancel Butttons
   'MsgBox("Do you like blue color?",3,"Choose options")

   ' Assume that you press No Button
   'MsgBox ("The Value of a is " & a)
   Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to continue?"
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "MsgBox Demonstration"
Help = "DEMO.HLP"
Ctxt = 1000
 Response = MsgBox("Do you like blue color ", 3, Title)
 'MsgBox Response
Global_Return = Response
MessageBox_Demo = Response
Rem Response = MsgBox(Msg, Style, Title, Help, Ctxt)
'If Response = vbYes Then    ' User chose Yes.
 '   MyString = "Yes"    ' Perform some action.
'Else    ' User chose No.
  '  MyString = "No"    ' Perform some action.
'End If
End Function
