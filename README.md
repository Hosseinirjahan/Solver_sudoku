Private Sub Command1_Click()
Option1.Value = False
Timer0.Interval = 1
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
For a = 1 To 81
Text1(a) = ""
Text1(a).BackColor = vbWhite
Next
For a = 1 To 9
Command7(a).Visible = True
Next
End Sub
Private Sub Command4_Click()
Option1.Value = True
End Sub
Private Sub Command5_Click()
For f = 1 To 81
Text2(f).Text = Text1(f).Text
Text2(f).ForeColor = Text1(f).ForeColor
Text2(f).BackColor = Text1(f).BackColor
Next
End Sub
Private Sub Command6_Click()
For a = 1 To 81
Text1(a).Text = Text2(a).Text
Text1(a).ForeColor = Text2(a).ForeColor
Text1(a).BackColor = Text2(a).BackColor
Next
End Sub
Private Sub Command7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y 
As Single)
b = Command7(Index).HelpContextID
If b <> 0 Then
Text1(b) = Index
Text1(b).ForeColor = vbBlack
Else
Text1(1) = Index
Text1(1).ForeColor = vbBlack
End If
End Sub
Private Sub Command8_Click()
MsgBox "ebteda jadval ro takmil konid baead az an ((Start)) bezanid ta jadval ra be soorat khodkar 
farakhani shavad.. "
MsgBox "shoma mitavanid safhe jadval ra zakhireh konid ba dokmehe ((Save)) va ba zadan dokmehe 
((Load)) va jadval zakhireh shodeh bazgardanid."
MsgBox "ba zadan dokmehe ((clear all)) shoma mitavanid tamam safheh ra pak konid; va agar khastid 
yek khaneh pak shavad har kelidi be joz adadhaye (1 ta 9) pak mishavad. "
MsgBox "zamani keh shoma dokmehs (Stop)) ra bezanid barnamehe az halateh test khodkar biroon 
miayad.."
MsgBox "zamani keh sodoko ta had emkan barnameh ra barayetan hal mikonad alamat zard cheshmak 
zan miavard ta 2 adad ra vared konid,keh shoma narm afzar ra zakhireh kardeh baed entekhab adad 
konid"
End Sub
Private Sub Text1_Change(Index As Integer)
If Text1(Index) <> "" Then
b = Asc(Text1(Index))
If b < 49 Or b > 57 Then
Text1(Index) = Chr(0)
End If
Else
Text1(Index).BackColor = vbWhite
End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Text1(Index) = Chr(0)
Text1(Index).ForeColor = vbBlack
End Sub
Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As 
Single)
Dim d As Integer
For e = 1 To 9
Label2(e).Caption = 0
Next e
For a = 1 To 81
Text1(a).HelpContextID = 0
For d = 1 To 9
If Text1(a).Text <> "" Then
 If Int(Text1(a).Text) = d Then
 Label2(d).Caption = Int(Label2(d).Caption) + 1
 End If
End If
Command7(d).Visible = True
Next d
Next a
b = Index
c = 0
If b >= 1 And b <= 9 Then
 For d = 1 To 9
 Text1(d).HelpContextID = 1
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
If b >= 10 And b <= 18 Then
 c = -18
 For d = 10 To 18
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
 
End If
If b >= 19 And b <= 27 Then
 c = -27
 For d = 19 To 27
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
 
End If
If b >= 28 And b <= 36 Then
 c = -36
 For d = 28 To 36
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
If b >= 37 And b <= 45 Then
 c = -45
 For d = 37 To 45
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
If b >= 46 And b <= 54 Then
 c = -54
 For d = 46 To 54
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
If b >= 55 And b <= 63 Then
 c = -63
 For d = 55 To 63
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
If b >= 64 And b <= 72 Then
 c = -72
 For d = 64 To 72
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
If b >= 73 And b <= 81 Then
 c = -81
 For d = 73 To 81
 Text1(d).HelpContextID = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).HelpContextID = 1
 End If
 
 Next d
End If
'bala samt chap
If (b >= 1 And b <= 3) Or (b >= 10 And b <= 12) Or (b >= 19 And b <= 21) Then
Text1(1).HelpContextID = 1
Text1(2).HelpContextID = 1
Text1(3).HelpContextID = 1
Text1(10).HelpContextID = 1
Text1(11).HelpContextID = 1
Text1(12).HelpContextID = 1
Text1(19).HelpContextID = 1
Text1(20).HelpContextID = 1
Text1(21).HelpContextID = 1
 
End If
'bala
If (b >= 4 And b <= 6) Or (b >= 13 And b <= 15) Or (b >= 22 And b <= 24) Then
Text1(4).HelpContextID = 1
Text1(5).HelpContextID = 1
Text1(6).HelpContextID = 1
Text1(13).HelpContextID = 1
Text1(14).HelpContextID = 1
Text1(15).HelpContextID = 1
Text1(22).HelpContextID = 1
Text1(23).HelpContextID = 1
Text1(24).HelpContextID = 1
End If
'bala samt rast
If (b >= 7 And b <= 9) Or (b >= 16 And b <= 18) Or (b >= 25 And b <= 27) Then
Text1(7).HelpContextID = 1
Text1(8).HelpContextID = 1
Text1(9).HelpContextID = 1
Text1(16).HelpContextID = 1
Text1(17).HelpContextID = 1
Text1(18).HelpContextID = 1
Text1(25).HelpContextID = 1
Text1(26).HelpContextID = 1
Text1(27).HelpContextID = 1
End If
'chap
If (b >= 28 And b <= 30) Or (b >= 37 And b <= 39) Or (b >= 46 And b <= 48) Then
Text1(28).HelpContextID = 1
Text1(29).HelpContextID = 1
Text1(30).HelpContextID = 1
Text1(37).HelpContextID = 1
Text1(38).HelpContextID = 1
Text1(39).HelpContextID = 1
Text1(46).HelpContextID = 1
Text1(47).HelpContextID = 1
Text1(48).HelpContextID = 1
End If
'vasat
If (b >= 31 And b <= 33) Or (b >= 40 And b <= 42) Or (b >= 49 And b <= 51) Then
Text1(31).HelpContextID = 1
Text1(32).HelpContextID = 1
Text1(33).HelpContextID = 1
Text1(40).HelpContextID = 1
Text1(41).HelpContextID = 1
Text1(42).HelpContextID = 1
Text1(49).HelpContextID = 1
Text1(50).HelpContextID = 1
Text1(51).HelpContextID = 1
End If
'rast
If (b >= 34 And b <= 36) Or (b >= 43 And b <= 45) Or (b >= 52 And b <= 54) Then
Text1(34).HelpContextID = 1
Text1(35).HelpContextID = 1
Text1(36).HelpContextID = 1
Text1(43).HelpContextID = 1
Text1(44).HelpContextID = 1
Text1(45).HelpContextID = 1
Text1(52).HelpContextID = 1
Text1(53).HelpContextID = 1
Text1(54).HelpContextID = 1
End If
'paean samt chap
If (b >= 55 And b <= 57) Or (b >= 64 And b <= 66) Or (b >= 73 And b <= 75) Then
Text1(55).HelpContextID = 1
Text1(56).HelpContextID = 1
Text1(57).HelpContextID = 1
Text1(64).HelpContextID = 1
Text1(65).HelpContextID = 1
Text1(66).HelpContextID = 1
Text1(73).HelpContextID = 1
Text1(74).HelpContextID = 1
Text1(75).HelpContextID = 1
End If
'paean
If (b >= 58 And b <= 60) Or (b >= 67 And b <= 69) Or (b >= 76 And b <= 78) Then
Text1(58).HelpContextID = 1
Text1(59).HelpContextID = 1
Text1(60).HelpContextID = 1
Text1(67).HelpContextID = 1
Text1(68).HelpContextID = 1
Text1(69).HelpContextID = 1
Text1(76).HelpContextID = 1
Text1(77).HelpContextID = 1
Text1(78).HelpContextID = 1
End If
'paean samt rast
If (b >= 61 And b <= 63) Or (b >= 70 And b <= 72) Or (b >= 79 And b <= 81) Then
Text1(61).HelpContextID = 1
Text1(62).HelpContextID = 1
Text1(63).HelpContextID = 1
Text1(70).HelpContextID = 1
Text1(71).HelpContextID = 1
Text1(72).HelpContextID = 1
Text1(79).HelpContextID = 1
Text1(80).HelpContextID = 1
Text1(81).HelpContextID = 1
End If
For a = 1 To 9
For b = 1 To 81
If Text1(b).HelpContextID = 1 Then
If Text1(b) = a Then
Command7(a).Visible = False
End If
End If
Next b
Command7(a).HelpContextID = Index
Next a
If Text1(Index) <> "" Then
For e = 1 To 9
Command7(e).Visible = False
Next e
End If
End Sub
Private Sub Timer0_Timer()
'barayeh fahmidan eshtebah radif 1
For a = 1 To 9
For b = 1 To 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 1 To 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 2
For a = 1 To 9
For b = 10 To 18
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 10 To 18
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 3
For a = 1 To 9
For b = 19 To 27
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 19 To 27
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 4
For a = 1 To 9
For b = 28 To 36
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 28 To 36
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 5
For a = 1 To 9
For b = 37 To 45
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 37 To 45
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 6
For a = 1 To 9
For b = 46 To 54
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 46 To 54
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 7
For a = 1 To 9
For b = 55 To 63
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 55 To 63
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 8
For a = 1 To 9
For b = 64 To 72
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 64 To 72
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah radif 9
For a = 1 To 9
For b = 73 To 81
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
 
For d = 73 To 81
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 1
For a = 1 To 9
For b = 1 To 73 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 1 To 73 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 2
For a = 1 To 9
For b = 2 To 74 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 2 To 74 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 3
For a = 1 To 9
For b = 3 To 75 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 3 To 75 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 4
For a = 1 To 9
For b = 4 To 76 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 4 To 76 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 5
For a = 1 To 9
For b = 5 To 77 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 5 To 77 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 6
For a = 1 To 9
For b = 6 To 78 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 6 To 78 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 7
For a = 1 To 9
For b = 7 To 79 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 7 To 79 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 8
For a = 1 To 9
For b = 8 To 80 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 8 To 80 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah setoon 9
For a = 1 To 9
For b = 9 To 81 Step 9
If Text1(b) = a Then
c = c + 1
End If
If c >= 2 Then
For d = 9 To 81 Step 9
Text1(d).BackColor = vbRed
Next d
End If
Next
c = 0
Next
'barayeh fahmidan eshtebah bala samt chap
 For a = 1 To 9
 For b = 1 To 21
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 1 To 21
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 
'barayeh fahmidan eshtebah bala samt vasat
 For a = 1 To 9
 For b = 4 To 24
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 4 To 24
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 
 'barayeh fahmidan eshtebah bala samt rast
 For a = 1 To 9
 For b = 7 To 27
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 7 To 27
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 'barayeh fahmidan eshtebah chap samt chap
 For a = 1 To 9
 For b = 28 To 48
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 28 To 48
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 'barayeh fahmidan eshtebah vasat samt vasat
 For a = 1 To 9
 For b = 31 To 51
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 31 To 51
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 'barayeh fahmidan eshtebah rast samt rast
 For a = 1 To 9
 For b = 34 To 54
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 34 To 54
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 'barayeh fahmidan eshtebah paean samt chap
 For a = 1 To 9
 For b = 55 To 75
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 55 To 75
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 'barayeh fahmidan eshtebah paean samt paean
 For a = 1 To 9
 For b = 58 To 78
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 58 To 78
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
 'barayeh fahmidan eshtebah paean samt rast
 For a = 1 To 9
 For b = 61 To 81
 If Text1(b) = a Then
 c = c + 1
 End If
 
 If c >= 2 Then
 For d = 61 To 81
 Text1(d).BackColor = vbRed
 
 If d Mod 3 = 0 Then
 d = d + 6
 End If
 Next d
 End If
 If b Mod 3 = 0 Then
 b = b + 6
 End If
 Next
 c = 0
 Next
For d = 1 To 81
If Text1(d).BackColor = vbRed Then
Timer0.Interval = 0
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Exit Sub
End If
Next d
Timer0.Interval = 0
If Option1.Value = False Then
Timer1.Interval = 1
End If
End Sub
Private Sub Timer1_Timer()
For a = 1 To 81
If Text1(a) = "" Then
Text1(a).BackColor = vbWhite
End If
Next
For a = 1 To 9
For b = 1 To 81
c = 0
If Text1(b) = a Then
If b >= 1 And b <= 9 Then
 For d = 1 To 9
 Text1(d).BackColor = vbGreen
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 10 And b <= 18 Then
 c = -18
 For d = 10 To 18
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
 
End If
If b >= 19 And b <= 27 Then
 c = -27
 For d = 19 To 27
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
 
End If
If b >= 28 And b <= 36 Then
 c = -36
 For d = 28 To 36
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 37 And b <= 45 Then
 c = -45
 For d = 37 To 45
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 46 And b <= 54 Then
 c = -54
 For d = 46 To 54
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 55 And b <= 63 Then
 c = -63
 For d = 55 To 63
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 64 And b <= 72 Then
 c = -72
 For d = 64 To 72
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 73 And b <= 81 Then
 c = -81
 For d = 73 To 81
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
'bala samt chap
If (b >= 1 And b <= 3) Or (b >= 10 And b <= 12) Or (b >= 19 And b <= 21) Then
Text1(1).BackColor = vbGreen
Text1(2).BackColor = vbGreen
Text1(3).BackColor = vbGreen
Text1(10).BackColor = vbGreen
Text1(11).BackColor = vbGreen
Text1(12).BackColor = vbGreen
Text1(19).BackColor = vbGreen
Text1(20).BackColor = vbGreen
Text1(21).BackColor = vbGreen
 
End If
'bala
If (b >= 4 And b <= 6) Or (b >= 13 And b <= 15) Or (b >= 22 And b <= 24) Then
Text1(4).BackColor = vbGreen
Text1(5).BackColor = vbGreen
Text1(6).BackColor = vbGreen
Text1(13).BackColor = vbGreen
Text1(14).BackColor = vbGreen
Text1(15).BackColor = vbGreen
Text1(22).BackColor = vbGreen
Text1(23).BackColor = vbGreen
Text1(24).BackColor = vbGreen
End If
'bala samt rast
If (b >= 7 And b <= 9) Or (b >= 16 And b <= 18) Or (b >= 25 And b <= 27) Then
Text1(7).BackColor = vbGreen
Text1(8).BackColor = vbGreen
Text1(9).BackColor = vbGreen
Text1(16).BackColor = vbGreen
Text1(17).BackColor = vbGreen
Text1(18).BackColor = vbGreen
Text1(25).BackColor = vbGreen
Text1(26).BackColor = vbGreen
Text1(27).BackColor = vbGreen
End If
'chap
If (b >= 28 And b <= 30) Or (b >= 37 And b <= 39) Or (b >= 46 And b <= 48) Then
Text1(28).BackColor = vbGreen
Text1(29).BackColor = vbGreen
Text1(30).BackColor = vbGreen
Text1(37).BackColor = vbGreen
Text1(38).BackColor = vbGreen
Text1(39).BackColor = vbGreen
Text1(46).BackColor = vbGreen
Text1(47).BackColor = vbGreen
Text1(48).BackColor = vbGreen
End If
'vasat
If (b >= 31 And b <= 33) Or (b >= 40 And b <= 42) Or (b >= 49 And b <= 51) Then
Text1(31).BackColor = vbGreen
Text1(32).BackColor = vbGreen
Text1(33).BackColor = vbGreen
Text1(40).BackColor = vbGreen
Text1(41).BackColor = vbGreen
Text1(42).BackColor = vbGreen
Text1(49).BackColor = vbGreen
Text1(50).BackColor = vbGreen
Text1(51).BackColor = vbGreen
End If
'rast
If (b >= 34 And b <= 36) Or (b >= 43 And b <= 45) Or (b >= 52 And b <= 54) Then
Text1(34).BackColor = vbGreen
Text1(35).BackColor = vbGreen
Text1(36).BackColor = vbGreen
Text1(43).BackColor = vbGreen
Text1(44).BackColor = vbGreen
Text1(45).BackColor = vbGreen
Text1(52).BackColor = vbGreen
Text1(53).BackColor = vbGreen
Text1(54).BackColor = vbGreen
End If
'paean samt chap
If (b >= 55 And b <= 57) Or (b >= 64 And b <= 66) Or (b >= 73 And b <= 75) Then
Text1(55).BackColor = vbGreen
Text1(56).BackColor = vbGreen
Text1(57).BackColor = vbGreen
Text1(64).BackColor = vbGreen
Text1(65).BackColor = vbGreen
Text1(66).BackColor = vbGreen
Text1(73).BackColor = vbGreen
Text1(74).BackColor = vbGreen
Text1(75).BackColor = vbGreen
End If
'paean
If (b >= 58 And b <= 60) Or (b >= 67 And b <= 69) Or (b >= 76 And b <= 78) Then
Text1(58).BackColor = vbGreen
Text1(59).BackColor = vbGreen
Text1(60).BackColor = vbGreen
Text1(67).BackColor = vbGreen
Text1(68).BackColor = vbGreen
Text1(69).BackColor = vbGreen
Text1(76).BackColor = vbGreen
Text1(77).BackColor = vbGreen
Text1(78).BackColor = vbGreen
End If
'paean samt rast
If (b >= 61 And b <= 63) Or (b >= 70 And b <= 72) Or (b >= 79 And b <= 81) Then
Text1(61).BackColor = vbGreen
Text1(62).BackColor = vbGreen
Text1(63).BackColor = vbGreen
Text1(70).BackColor = vbGreen
Text1(71).BackColor = vbGreen
Text1(72).BackColor = vbGreen
Text1(79).BackColor = vbGreen
Text1(80).BackColor = vbGreen
Text1(81).BackColor = vbGreen
End If
End If
Next b
'bala samt chap
If (Text1(1).BackColor = vbWhite And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(1) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbWhite And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(2) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbWhite) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(3) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbWhite And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(10) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbWhite And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(11) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbWhite) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(12) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbWhite And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbGreen) Then
Text1(19) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbWhite And 
Text1(21).BackColor = vbGreen) Then
Text1(20) = a
End If
If (Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen) 
And (Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen) And (Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And 
Text1(21).BackColor = vbWhite) Then
Text1(21) = a
End If
'bala
If (Text1(4).BackColor = vbWhite And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen) Then
Text1(4) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbWhite And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen) Then
Text1(5) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbWhite) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen) Then
Text1(6) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbWhite And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen) Then
Text1(13) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen)
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbWhite And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen) Then
Text1(14) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbWhite) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And
Text1(24).BackColor = vbGreen) Then
Text1(15) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbWhite And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen) Then
Text1(22) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbWhite And 
Text1(24).BackColor = vbGreen) Then
Text1(23) = a
End If
If (Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen) 
And (Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(15).BackColor = 
vbGreen) And (Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbWhite) Then
Text1(24) = a
End If
'bala samt rast
If (Text1(7).BackColor = vbWhite And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(7) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbWhite And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(8) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbWhite) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(9) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbWhite And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(16) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbWhite And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(17) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbWhite) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(18) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbWhite And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbGreen) Then
Text1(25) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbWhite And 
Text1(27).BackColor = vbGreen) Then
Text1(26) = a
End If
If (Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen) 
And (Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(18).BackColor = 
vbGreen) And (Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen And 
Text1(27).BackColor = vbWhite) Then
Text1(27) = a
End If
'chap
If (Text1(28).BackColor = vbWhite And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(28) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbWhite And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(29) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbWhite) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(30) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbWhite And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(37) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbWhite And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(38) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbWhite) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(39) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbWhite And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbGreen) Then
Text1(46) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbWhite And Text1(48).BackColor = vbGreen) Then
Text1(47) = a
End If
If (Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen) And (Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And 
Text1(39).BackColor = vbGreen) And (Text1(46).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(48).BackColor = vbWhite) Then
Text1(48) = a
End If
'vasat
If (Text1(31).BackColor = vbWhite And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(31) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbWhite And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(32) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbWhite) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(33) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbWhite And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(40) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbWhite And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(41) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbWhite) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(42) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbWhite And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbGreen) Then
Text1(49) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbWhite And Text1(51).BackColor = vbGreen) Then
Text1(50) = a
End If
If (Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And Text1(33).BackColor = 
vbGreen) And (Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen) And (Text1(49).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(51).BackColor = vbWhite) Then
Text1(51) = a
End If
'rast
If (Text1(34).BackColor = vbWhite And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(34) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbWhite And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(35) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbWhite) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(36) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbWhite And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(43) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbWhite And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(44) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbWhite) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(45) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbWhite And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbGreen) Then
Text1(52) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbWhite And Text1(54).BackColor = vbGreen) Then
Text1(53) = a
End If
If (Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen And Text1(36).BackColor = 
vbGreen) And (Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen And 
Text1(45).BackColor = vbGreen) And (Text1(52).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(54).BackColor = vbWhite) Then
Text1(54) = a
End If
'paean samt chap
If (Text1(55).BackColor = vbWhite And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(55) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbWhite And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(56) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbWhite) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(57) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbWhite And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(64) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbWhite And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(65) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbWhite) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(66) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbWhite And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbGreen) Then
Text1(73) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbWhite And Text1(75).BackColor = vbGreen) Then
Text1(74) = a
End If
If (Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen) And (Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(66).BackColor = vbGreen) And (Text1(73).BackColor = vbGreen And Text1(74).BackColor = 
vbGreen And Text1(75).BackColor = vbWhite) Then
Text1(75) = a
End If
'paean
If (Text1(58).BackColor = vbWhite And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(58) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbWhite And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(59) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbWhite) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(60) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbWhite And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(67) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbWhite And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(68) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbWhite) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(69) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbWhite And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbGreen) Then
Text1(76) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbWhite And Text1(78).BackColor = vbGreen) Then
Text1(77) = a
End If
If (Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And Text1(60).BackColor = 
vbGreen) And (Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen) And (Text1(76).BackColor = vbGreen And Text1(77).BackColor = 
vbGreen And Text1(78).BackColor = vbWhite) Then
Text1(78) = a
End If
'paean samt rast
If (Text1(61).BackColor = vbWhite And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(61) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbWhite And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(62) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbWhite) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(63) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbWhite And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(70) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbWhite And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(71) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbWhite) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(72) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbWhite And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbGreen) Then
Text1(79) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbWhite And Text1(81).BackColor = vbGreen) Then
Text1(80) = a
End If
If (Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen And Text1(63).BackColor = 
vbGreen) And (Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(72).BackColor = vbGreen) And (Text1(79).BackColor = vbGreen And Text1(80).BackColor = 
vbGreen And Text1(81).BackColor = vbWhite) Then
Text1(81) = a
End If
'baray moshakhas kardan adadha
For e = 1 To 81
If Text1(e) = "" Then
Text1(e).BackColor = vbWhite
End If
Next e
Next a
Timer1.Interval = 0
If Option1.Value = False Then
Timer5.Interval = 1
End If
End Sub
Private Sub Timer2_Timer()
For a = 1 To 81
If Text1(a) = "" Then
Text1(a).BackColor = vbWhite
Else
Text1(a).BackColor = vbGreen
End If
Next a
'baraye moghayesehe radif
'radif avval
If Text1(1).BackColor = vbWhite And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(1) = 45 - (Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + Val(Text1(5).Text) + 
Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbWhite And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(2) = 45 - (Val(Text1(1).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + Val(Text1(5).Text) + 
Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbWhite 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(3) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(4).Text) + Val(Text1(5).Text) + 
Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbWhite And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(4) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(5).Text) + 
Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbWhite And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(5) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + 
Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbWhite 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(6) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + 
Val(Text1(5).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbWhite And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(7) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + 
Val(Text1(5).Text) + Val(Text1(6).Text) + Val(Text1(8).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbWhite And Text1(9).BackColor = vbGreen 
Then
Text1(8) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + 
Val(Text1(5).Text) + Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(9).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbWhite 
Then
Text1(9) = 45 - (Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text) + Val(Text1(4).Text) + 
Val(Text1(5).Text) + Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text))
End If
'radif dovvom
If Text1(10).BackColor = vbWhite And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(10) = 45 - (Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + Val(Text1(14).Text) + 
Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbWhite And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(11) = 45 - (Val(Text1(10).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + Val(Text1(14).Text) + 
Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbWhite And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(12) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(13).Text) + Val(Text1(14).Text) + 
Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbWhite And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(13) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(14).Text) + 
Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbWhite And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(14) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + 
Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbWhite And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(15) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + 
Val(Text1(14).Text) + Val(Text1(16).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbWhite And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(16) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + 
Val(Text1(14).Text) + Val(Text1(15).Text) + Val(Text1(17).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbWhite 
And Text1(18).BackColor = vbGreen Then
Text1(17) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + 
Val(Text1(14).Text) + Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(18).Text))
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbWhite Then
Text1(18) = 45 - (Val(Text1(10).Text) + Val(Text1(11).Text) + Val(Text1(12).Text) + Val(Text1(13).Text) + 
Val(Text1(14).Text) + Val(Text1(15).Text) + Val(Text1(16).Text) + Val(Text1(17).Text))
End If
'radif sevvom
If Text1(19).BackColor = vbWhite And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(19) = 45 - (Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + Val(Text1(23).Text) + 
Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbWhite And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(20) = 45 - (Val(Text1(19).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + Val(Text1(23).Text) + 
Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbWhite And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(21) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(22).Text) + Val(Text1(23).Text) + 
Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbWhite And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(22) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(23).Text) + 
Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbWhite And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(23) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + 
Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbWhite And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(24) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + 
Val(Text1(23).Text) + Val(Text1(25).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbWhite And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(25) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + 
Val(Text1(23).Text) + Val(Text1(24).Text) + Val(Text1(26).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbWhite 
And Text1(27).BackColor = vbGreen Then
Text1(26) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + 
Val(Text1(23).Text) + Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(27).Text))
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbWhite Then
Text1(27) = 45 - (Val(Text1(19).Text) + Val(Text1(20).Text) + Val(Text1(21).Text) + Val(Text1(22).Text) + 
Val(Text1(23).Text) + Val(Text1(24).Text) + Val(Text1(25).Text) + Val(Text1(26).Text))
End If
'radif chaharom
If Text1(28).BackColor = vbWhite And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(28) = 45 - (Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + Val(Text1(32).Text) + 
Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbWhite And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(29) = 45 - (Val(Text1(28).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + Val(Text1(32).Text) + 
Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbWhite And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(30) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(31).Text) + Val(Text1(32).Text) + 
Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbWhite And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(31) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(32).Text) + 
Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbWhite And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(32) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + 
Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbWhite And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(33) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + 
Val(Text1(32).Text) + Val(Text1(34).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbWhite And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(34) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + 
Val(Text1(32).Text) + Val(Text1(33).Text) + Val(Text1(35).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbWhite 
And Text1(36).BackColor = vbGreen Then
Text1(35) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + 
Val(Text1(32).Text) + Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(36).Text))
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbWhite Then
Text1(36) = 45 - (Val(Text1(28).Text) + Val(Text1(29).Text) + Val(Text1(30).Text) + Val(Text1(31).Text) + 
Val(Text1(32).Text) + Val(Text1(33).Text) + Val(Text1(34).Text) + Val(Text1(35).Text))
End If
'radif panjom
If Text1(37).BackColor = vbWhite And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(37) = 45 - (Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + Val(Text1(41).Text) + 
Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbWhite And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(38) = 45 - (Val(Text1(37).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + Val(Text1(41).Text) + 
Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbWhite And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(39) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(40).Text) + Val(Text1(41).Text) + 
Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbWhite And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(40) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(41).Text) + 
Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbWhite And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(41) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + 
Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbWhite And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(42) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + 
Val(Text1(41).Text) + Val(Text1(43).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbWhite And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(43) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + 
Val(Text1(41).Text) + Val(Text1(42).Text) + Val(Text1(44).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbWhite 
And Text1(45).BackColor = vbGreen Then
Text1(44) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + 
Val(Text1(41).Text) + Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(45).Text))
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbWhite Then
Text1(45) = 45 - (Val(Text1(37).Text) + Val(Text1(38).Text) + Val(Text1(39).Text) + Val(Text1(40).Text) + 
Val(Text1(41).Text) + Val(Text1(42).Text) + Val(Text1(43).Text) + Val(Text1(44).Text))
End If
'radif sheshom
If Text1(46).BackColor = vbWhite And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(46) = 45 - (Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + Val(Text1(50).Text) + 
Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbWhite And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(47) = 45 - (Val(Text1(46).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + Val(Text1(50).Text) + 
Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbWhite And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(48) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(49).Text) + Val(Text1(50).Text) + 
Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbWhite And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(49) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(50).Text) + 
Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbWhite And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(50) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + 
Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbWhite And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(51) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + 
Val(Text1(50).Text) + Val(Text1(52).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbWhite And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(52) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + 
Val(Text1(50).Text) + Val(Text1(51).Text) + Val(Text1(53).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbWhite 
And Text1(54).BackColor = vbGreen Then
Text1(53) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + 
Val(Text1(50).Text) + Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(54).Text))
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbWhite Then
Text1(54) = 45 - (Val(Text1(46).Text) + Val(Text1(47).Text) + Val(Text1(48).Text) + Val(Text1(49).Text) + 
Val(Text1(50).Text) + Val(Text1(51).Text) + Val(Text1(52).Text) + Val(Text1(53).Text))
End If
'radif haftom
If Text1(55).BackColor = vbWhite And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(55) = 45 - (Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + Val(Text1(59).Text) + 
Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbWhite And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(56) = 45 - (Val(Text1(55).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + Val(Text1(59).Text) + 
Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbWhite And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(57) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(58).Text) + Val(Text1(59).Text) + 
Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbWhite And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(58) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(59).Text) + 
Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbWhite And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(59) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + 
Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbWhite And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(60) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + 
Val(Text1(59).Text) + Val(Text1(61).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbWhite And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(61) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + 
Val(Text1(59).Text) + Val(Text1(60).Text) + Val(Text1(62).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbWhite 
And Text1(63).BackColor = vbGreen Then
Text1(62) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + 
Val(Text1(59).Text) + Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(63).Text))
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbWhite Then
Text1(63) = 45 - (Val(Text1(55).Text) + Val(Text1(56).Text) + Val(Text1(57).Text) + Val(Text1(58).Text) + 
Val(Text1(59).Text) + Val(Text1(60).Text) + Val(Text1(61).Text) + Val(Text1(62).Text))
End If
Timer2.Interval = 0
If Option1.Value = False Then
Timer3.Interval = 1
End If
End Sub
Private Sub Timer3_Timer()
'radif hashtom
If Text1(64).BackColor = vbWhite And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(64) = 45 - (Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + Val(Text1(68).Text) + 
Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbWhite And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(65) = 45 - (Val(Text1(64).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + Val(Text1(68).Text) + 
Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbWhite And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(66) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(67).Text) + Val(Text1(68).Text) + 
Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbWhite And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(67) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(68).Text) + 
Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbWhite And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(68) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + 
Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbWhite And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(69) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + 
Val(Text1(68).Text) + Val(Text1(70).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbWhite And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(70) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + 
Val(Text1(68).Text) + Val(Text1(69).Text) + Val(Text1(71).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbWhite 
And Text1(72).BackColor = vbGreen Then
Text1(71) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + 
Val(Text1(68).Text) + Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(72).Text))
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbWhite Then
Text1(72) = 45 - (Val(Text1(64).Text) + Val(Text1(65).Text) + Val(Text1(66).Text) + Val(Text1(67).Text) + 
Val(Text1(68).Text) + Val(Text1(69).Text) + Val(Text1(70).Text) + Val(Text1(71).Text))
End If
'radif nohom
If Text1(73).BackColor = vbWhite And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(73) = 45 - (Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + Val(Text1(77).Text) + 
Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbWhite And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(74) = 45 - (Val(Text1(73).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + Val(Text1(77).Text) + 
Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbWhite And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(75) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(76).Text) + Val(Text1(77).Text) + 
Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbWhite And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(76) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(77).Text) + 
Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbWhite And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(77) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + 
Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbWhite And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(78) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + 
Val(Text1(77).Text) + Val(Text1(79).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbWhite And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(79) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + 
Val(Text1(77).Text) + Val(Text1(78).Text) + Val(Text1(80).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbWhite 
And Text1(81).BackColor = vbGreen Then
Text1(80) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + 
Val(Text1(77).Text) + Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(81).Text))
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbWhite Then
Text1(81) = 45 - (Val(Text1(73).Text) + Val(Text1(74).Text) + Val(Text1(75).Text) + Val(Text1(76).Text) + 
Val(Text1(77).Text) + Val(Text1(78).Text) + Val(Text1(79).Text) + Val(Text1(80).Text))
End If
'baraye moghayesehe setoon
'setoon avval
If Text1(1).BackColor = vbWhite And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(1) = 45 - (Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + Val(Text1(37).Text) + 
Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbWhite And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(10) = 45 - (Val(Text1(1).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + Val(Text1(37).Text) + 
Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbWhite 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(19) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(28).Text) + Val(Text1(37).Text) + 
Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbWhite And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(28) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(37).Text) + 
Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbWhite And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(37) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + 
Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbWhite And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(46) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + 
Val(Text1(37).Text) + Val(Text1(55).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbWhite And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(55) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + 
Val(Text1(37).Text) + Val(Text1(46).Text) + Val(Text1(64).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbWhite And 
Text1(73).BackColor = vbGreen Then
Text1(64) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + 
Val(Text1(37).Text) + Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(73).Text))
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbWhite Then
Text1(73) = 45 - (Val(Text1(1).Text) + Val(Text1(10).Text) + Val(Text1(19).Text) + Val(Text1(28).Text) + 
Val(Text1(37).Text) + Val(Text1(46).Text) + Val(Text1(55).Text) + Val(Text1(64).Text))
End If
'setoon dovvom
If Text1(2).BackColor = vbWhite And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(2) = 45 - (Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + Val(Text1(38).Text) + 
Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbWhite And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(11) = 45 - (Val(Text1(2).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + Val(Text1(38).Text) + 
Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbWhite 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(20) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(29).Text) + Val(Text1(38).Text) + 
Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbWhite And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(29) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(38).Text) + 
Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbWhite And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(38) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + 
Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbWhite And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(47) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + 
Val(Text1(38).Text) + Val(Text1(56).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbWhite And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(56) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + 
Val(Text1(38).Text) + Val(Text1(47).Text) + Val(Text1(65).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbWhite And 
Text1(74).BackColor = vbGreen Then
Text1(65) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + 
Val(Text1(38).Text) + Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(74).Text))
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbWhite Then
Text1(74) = 45 - (Val(Text1(2).Text) + Val(Text1(11).Text) + Val(Text1(20).Text) + Val(Text1(29).Text) + 
Val(Text1(38).Text) + Val(Text1(47).Text) + Val(Text1(56).Text) + Val(Text1(65).Text))
End If
'setoon sevvom
If Text1(3).BackColor = vbWhite And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(3) = 45 - (Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + Val(Text1(39).Text) + 
Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbWhite And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(12) = 45 - (Val(Text1(3).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + Val(Text1(39).Text) + 
Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbWhite 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(21) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(30).Text) + Val(Text1(39).Text) + 
Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbWhite 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(21) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(30).Text) + Val(Text1(39).Text) + 
Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbWhite And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(30) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(39).Text) + 
Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbWhite And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(39) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + 
Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbWhite And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(48) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + 
Val(Text1(39).Text) + Val(Text1(57).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbWhite And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(57) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + 
Val(Text1(39).Text) + Val(Text1(48).Text) + Val(Text1(66).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbWhite And 
Text1(75).BackColor = vbGreen Then
Text1(66) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + 
Val(Text1(39).Text) + Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(75).Text))
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbWhite Then
Text1(75) = 45 - (Val(Text1(3).Text) + Val(Text1(12).Text) + Val(Text1(21).Text) + Val(Text1(30).Text) + 
Val(Text1(39).Text) + Val(Text1(48).Text) + Val(Text1(57).Text) + Val(Text1(66).Text))
End If
'setton charom
If Text1(4).BackColor = vbWhite And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(4) = 45 - (Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + Val(Text1(40).Text) + 
Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbWhite And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(13) = 45 - (Val(Text1(4).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + Val(Text1(40).Text) + 
Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbWhite 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(22) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(31).Text) + Val(Text1(40).Text) + 
Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbWhite And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(31) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(40).Text) + 
Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbWhite And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(40) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + 
Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbWhite And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(49) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + 
Val(Text1(40).Text) + Val(Text1(58).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbWhite And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(58) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + 
Val(Text1(40).Text) + Val(Text1(49).Text) + Val(Text1(67).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbWhite And 
Text1(76).BackColor = vbGreen Then
Text1(67) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + 
Val(Text1(40).Text) + Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(76).Text))
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbWhite Then
Text1(76) = 45 - (Val(Text1(4).Text) + Val(Text1(13).Text) + Val(Text1(22).Text) + Val(Text1(31).Text) + 
Val(Text1(40).Text) + Val(Text1(49).Text) + Val(Text1(58).Text) + Val(Text1(67).Text))
End If
'setoon panjom
If Text1(5).BackColor = vbWhite And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(5) = 45 - (Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + Val(Text1(41).Text) + 
Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbWhite And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(14) = 45 - (Val(Text1(5).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + Val(Text1(41).Text) + 
Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbWhite 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(23) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(32).Text) + Val(Text1(41).Text) + 
Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbWhite And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(32) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(41).Text) + 
Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbWhite And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(41) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + 
Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbWhite And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(50) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + 
Val(Text1(41).Text) + Val(Text1(59).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbWhite And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(59) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + 
Val(Text1(41).Text) + Val(Text1(50).Text) + Val(Text1(68).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbWhite And 
Text1(77).BackColor = vbGreen Then
Text1(68) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + 
Val(Text1(41).Text) + Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(77).Text))
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbWhite Then
Text1(77) = 45 - (Val(Text1(5).Text) + Val(Text1(14).Text) + Val(Text1(23).Text) + Val(Text1(32).Text) + 
Val(Text1(41).Text) + Val(Text1(50).Text) + Val(Text1(59).Text) + Val(Text1(68).Text))
End If
'setoon sheshom
If Text1(6).BackColor = vbWhite And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(6) = 45 - (Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + Val(Text1(42).Text) + 
Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbWhite And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(15) = 45 - (Val(Text1(6).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + Val(Text1(42).Text) + 
Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbWhite 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(24) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(33).Text) + Val(Text1(42).Text) + 
Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbWhite And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(33) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(42).Text) + 
Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbWhite And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(42) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + 
Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbWhite And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(51) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + 
Val(Text1(42).Text) + Val(Text1(60).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbWhite And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(60) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + 
Val(Text1(42).Text) + Val(Text1(51).Text) + Val(Text1(69).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbWhite And 
Text1(78).BackColor = vbGreen Then
Text1(69) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + 
Val(Text1(42).Text) + Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(78).Text))
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbWhite Then
Text1(78) = 45 - (Val(Text1(6).Text) + Val(Text1(15).Text) + Val(Text1(24).Text) + Val(Text1(33).Text) + 
Val(Text1(42).Text) + Val(Text1(51).Text) + Val(Text1(60).Text) + Val(Text1(69).Text))
End If
Timer3.Interval = 0
If Option1.Value = False Then
Timer4.Interval = 1
End If
End Sub
Private Sub Timer4_Timer()
'setoon haftom
If Text1(7).BackColor = vbWhite And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(7) = 45 - (Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + Val(Text1(43).Text) + 
Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbWhite And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(16) = 45 - (Val(Text1(7).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + Val(Text1(43).Text) + 
Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbWhite 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(25) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(34).Text) + Val(Text1(43).Text) + 
Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbWhite And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(34) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(43).Text) + 
Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbWhite And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(43) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + 
Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbWhite And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(52) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + 
Val(Text1(43).Text) + Val(Text1(61).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbWhite And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(61) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + 
Val(Text1(43).Text) + Val(Text1(52).Text) + Val(Text1(70).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbWhite And 
Text1(79).BackColor = vbGreen Then
Text1(70) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + 
Val(Text1(43).Text) + Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(79).Text))
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbWhite Then
Text1(79) = 45 - (Val(Text1(7).Text) + Val(Text1(16).Text) + Val(Text1(25).Text) + Val(Text1(34).Text) + 
Val(Text1(43).Text) + Val(Text1(52).Text) + Val(Text1(61).Text) + Val(Text1(70).Text))
End If
'setoon hashtom
If Text1(8).BackColor = vbWhite And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(8) = 45 - (Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + Val(Text1(44).Text) + 
Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbWhite And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(17) = 45 - (Val(Text1(8).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + Val(Text1(44).Text) + 
Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbWhite 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(26) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(35).Text) + Val(Text1(44).Text) + 
Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbWhite And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(35) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(44).Text) + 
Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbWhite And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(44) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + 
Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbWhite And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(53) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + 
Val(Text1(44).Text) + Val(Text1(62).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbWhite And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(62) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + 
Val(Text1(44).Text) + Val(Text1(53).Text) + Val(Text1(71).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbWhite And 
Text1(80).BackColor = vbGreen Then
Text1(71) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + 
Val(Text1(44).Text) + Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(80).Text))
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbWhite Then
Text1(80) = 45 - (Val(Text1(8).Text) + Val(Text1(17).Text) + Val(Text1(26).Text) + Val(Text1(35).Text) + 
Val(Text1(44).Text) + Val(Text1(53).Text) + Val(Text1(62).Text) + Val(Text1(71).Text))
End If
'setoon nohom
If Text1(9).BackColor = vbWhite And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(9) = 45 - (Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + Val(Text1(45).Text) + 
Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbWhite And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(18) = 45 - (Val(Text1(9).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + Val(Text1(45).Text) + 
Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbWhite 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(27) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(36).Text) + Val(Text1(45).Text) + 
Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbWhite And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(36) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(45).Text) + 
Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbWhite And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(45) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + 
Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbWhite And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(54) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + 
Val(Text1(45).Text) + Val(Text1(63).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbWhite And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(63) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + 
Val(Text1(45).Text) + Val(Text1(54).Text) + Val(Text1(72).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbWhite And 
Text1(81).BackColor = vbGreen Then
Text1(72) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + 
Val(Text1(45).Text) + Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(81).Text))
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbWhite Then
Text1(81) = 45 - (Val(Text1(9).Text) + Val(Text1(18).Text) + Val(Text1(27).Text) + Val(Text1(36).Text) + 
Val(Text1(45).Text) + Val(Text1(54).Text) + Val(Text1(63).Text) + Val(Text1(72).Text))
End If
Timer4.Interval = 0
If Option1.Value = False Then
Timer5.Interval = 1
End If
End Sub
Private Sub Timer5_Timer()
For a = 1 To 81
If Text1(a) = "" Then
Text1(a).BackColor = vbWhite
End If
Next
For a = 1 To 9
For b = 1 To 81
c = 0
If Text1(b) = a Then
If b >= 1 And b <= 9 Then
 For d = 1 To 9
 Text1(d).BackColor = vbGreen
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 10 And b <= 18 Then
 c = -18
 For d = 10 To 18
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
 
End If
If b >= 19 And b <= 27 Then
 c = -27
 For d = 19 To 27
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
 
End If
If b >= 28 And b <= 36 Then
 c = -36
 For d = 28 To 36
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 37 And b <= 45 Then
 c = -45
 For d = 37 To 45
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 46 And b <= 54 Then
 c = -54
 For d = 46 To 54
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 55 And b <= 63 Then
 c = -63
 For d = 55 To 63
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 64 And b <= 72 Then
 c = -72
 For d = 64 To 72
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 73 And b <= 81 Then
 c = -81
 For d = 73 To 81
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
'bala samt chap
If (b >= 1 And b <= 3) Or (b >= 10 And b <= 12) Or (b >= 19 And b <= 21) Then
Text1(1).BackColor = vbGreen
Text1(2).BackColor = vbGreen
Text1(3).BackColor = vbGreen
Text1(10).BackColor = vbGreen
Text1(11).BackColor = vbGreen
Text1(12).BackColor = vbGreen
Text1(19).BackColor = vbGreen
Text1(20).BackColor = vbGreen
Text1(21).BackColor = vbGreen
 
End If
'bala
If (b >= 4 And b <= 6) Or (b >= 13 And b <= 15) Or (b >= 22 And b <= 24) Then
Text1(4).BackColor = vbGreen
Text1(5).BackColor = vbGreen
Text1(6).BackColor = vbGreen
Text1(13).BackColor = vbGreen
Text1(14).BackColor = vbGreen
Text1(15).BackColor = vbGreen
Text1(22).BackColor = vbGreen
Text1(23).BackColor = vbGreen
Text1(24).BackColor = vbGreen
End If
'bala samt rast
If (b >= 7 And b <= 9) Or (b >= 16 And b <= 18) Or (b >= 25 And b <= 27) Then
Text1(7).BackColor = vbGreen
Text1(8).BackColor = vbGreen
Text1(9).BackColor = vbGreen
Text1(16).BackColor = vbGreen
Text1(17).BackColor = vbGreen
Text1(18).BackColor = vbGreen
Text1(25).BackColor = vbGreen
Text1(26).BackColor = vbGreen
Text1(27).BackColor = vbGreen
End If
'chap
If (b >= 28 And b <= 30) Or (b >= 37 And b <= 39) Or (b >= 46 And b <= 48) Then
Text1(28).BackColor = vbGreen
Text1(29).BackColor = vbGreen
Text1(30).BackColor = vbGreen
Text1(37).BackColor = vbGreen
Text1(38).BackColor = vbGreen
Text1(39).BackColor = vbGreen
Text1(46).BackColor = vbGreen
Text1(47).BackColor = vbGreen
Text1(48).BackColor = vbGreen
End If
'vasat
If (b >= 31 And b <= 33) Or (b >= 40 And b <= 42) Or (b >= 49 And b <= 51) Then
Text1(31).BackColor = vbGreen
Text1(32).BackColor = vbGreen
Text1(33).BackColor = vbGreen
Text1(40).BackColor = vbGreen
Text1(41).BackColor = vbGreen
Text1(42).BackColor = vbGreen
Text1(49).BackColor = vbGreen
Text1(50).BackColor = vbGreen
Text1(51).BackColor = vbGreen
End If
'rast
If (b >= 34 And b <= 36) Or (b >= 43 And b <= 45) Or (b >= 52 And b <= 54) Then
Text1(34).BackColor = vbGreen
Text1(35).BackColor = vbGreen
Text1(36).BackColor = vbGreen
Text1(43).BackColor = vbGreen
Text1(44).BackColor = vbGreen
Text1(45).BackColor = vbGreen
Text1(52).BackColor = vbGreen
Text1(53).BackColor = vbGreen
Text1(54).BackColor = vbGreen
End If
'paean samt chap
If (b >= 55 And b <= 57) Or (b >= 64 And b <= 66) Or (b >= 73 And b <= 75) Then
Text1(55).BackColor = vbGreen
Text1(56).BackColor = vbGreen
Text1(57).BackColor = vbGreen
Text1(64).BackColor = vbGreen
Text1(65).BackColor = vbGreen
Text1(66).BackColor = vbGreen
Text1(73).BackColor = vbGreen
Text1(74).BackColor = vbGreen
Text1(75).BackColor = vbGreen
End If
'paean
If (b >= 58 And b <= 60) Or (b >= 67 And b <= 69) Or (b >= 76 And b <= 78) Then
Text1(58).BackColor = vbGreen
Text1(59).BackColor = vbGreen
Text1(60).BackColor = vbGreen
Text1(67).BackColor = vbGreen
Text1(68).BackColor = vbGreen
Text1(69).BackColor = vbGreen
Text1(76).BackColor = vbGreen
Text1(77).BackColor = vbGreen
Text1(78).BackColor = vbGreen
End If
'paean samt rast
If (b >= 61 And b <= 63) Or (b >= 70 And b <= 72) Or (b >= 79 And b <= 81) Then
Text1(61).BackColor = vbGreen
Text1(62).BackColor = vbGreen
Text1(63).BackColor = vbGreen
Text1(70).BackColor = vbGreen
Text1(71).BackColor = vbGreen
Text1(72).BackColor = vbGreen
Text1(79).BackColor = vbGreen
Text1(80).BackColor = vbGreen
Text1(81).BackColor = vbGreen
End If
End If
Next b
'baraye moghayesehe radif
'radif avval
If Text1(1).BackColor = vbWhite And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(1) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbWhite And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(2) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbWhite 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(3) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbWhite And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(4) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbWhite And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(5) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbWhite 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(6) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbWhite And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbGreen 
Then
Text1(7) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbWhite And Text1(9).BackColor = vbGreen 
Then
Text1(8) = a
End If
If Text1(1).BackColor = vbGreen And Text1(2).BackColor = vbGreen And Text1(3).BackColor = vbGreen 
And Text1(4).BackColor = vbGreen And Text1(5).BackColor = vbGreen And Text1(6).BackColor = vbGreen 
And Text1(7).BackColor = vbGreen And Text1(8).BackColor = vbGreen And Text1(9).BackColor = vbWhite 
Then
Text1(9) = a
End If
'radif dovvom
If Text1(10).BackColor = vbWhite And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(10) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbWhite And Text1(12).BackColor =
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(11) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbWhite And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(12) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbWhite And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(13) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbWhite And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(14) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbWhite And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(15) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbWhite And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbGreen Then
Text1(16) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbWhite 
And Text1(18).BackColor = vbGreen Then
Text1(17) = a
End If
If Text1(10).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(12).BackColor = 
vbGreen And Text1(13).BackColor = vbGreen And Text1(14).BackColor = vbGreen And 
Text1(15).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(17).BackColor = vbGreen 
And Text1(18).BackColor = vbWhite Then
Text1(18) = a
End If
'radif sevvom
If Text1(19).BackColor = vbWhite And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(19) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbWhite And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(20) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbWhite And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(21) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbWhite And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(22) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbWhite And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(23) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbWhite And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(24) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbWhite And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbGreen Then
Text1(25) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbWhite 
And Text1(27).BackColor = vbGreen Then
Text1(26) = a
End If
If Text1(19).BackColor = vbGreen And Text1(20).BackColor = vbGreen And Text1(21).BackColor = 
vbGreen And Text1(22).BackColor = vbGreen And Text1(23).BackColor = vbGreen And 
Text1(24).BackColor = vbGreen And Text1(25).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(27).BackColor = vbWhite Then
Text1(27) = a
End If
'radif chaharom
If Text1(28).BackColor = vbWhite And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(28) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbWhite And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(29) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbWhite And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(30) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbWhite And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(31) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbWhite And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(32) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbWhite And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(33) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbWhite And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen Then
Text1(34) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbWhite 
And Text1(36).BackColor = vbGreen Then
Text1(35) = a
End If
If Text1(28).BackColor = vbGreen And Text1(29).BackColor = vbGreen And Text1(30).BackColor = 
vbGreen And Text1(31).BackColor = vbGreen And Text1(32).BackColor = vbGreen And 
Text1(33).BackColor = vbGreen And Text1(34).BackColor = vbGreen And Text1(35).BackColor = vbGreen 
And Text1(36).BackColor = vbWhite Then
Text1(36) = a
End If
'radif panjom
If Text1(37).BackColor = vbWhite And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(37) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbWhite And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(38) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbWhite And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(39) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbWhite And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(40) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbWhite And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(41) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbWhite And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(42) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbWhite And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbGreen Then
Text1(43) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbWhite 
And Text1(45).BackColor = vbGreen Then
Text1(44) = a
End If
If Text1(37).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(39).BackColor = 
vbGreen And Text1(40).BackColor = vbGreen And Text1(41).BackColor = vbGreen And 
Text1(42).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(44).BackColor = vbGreen 
And Text1(45).BackColor = vbWhite Then
Text1(45) = a
End If
'radif sheshom
If Text1(46).BackColor = vbWhite And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(46) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbWhite And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(47) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbWhite And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(48) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbWhite And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(49) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbWhite And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(50) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbWhite And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(51) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbWhite And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbGreen Then
Text1(52) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbWhite 
And Text1(54).BackColor = vbGreen Then
Text1(53) = a
End If
If Text1(46).BackColor = vbGreen And Text1(47).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(49).BackColor = vbGreen And Text1(50).BackColor = vbGreen And 
Text1(51).BackColor = vbGreen And Text1(52).BackColor = vbGreen And Text1(53).BackColor = vbGreen 
And Text1(54).BackColor = vbWhite Then
Text1(54) = a
End If
'radif haftom
If Text1(55).BackColor = vbWhite And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(55) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbWhite And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(56) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbWhite And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen
And Text1(63).BackColor = vbGreen Then
Text1(57) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbWhite And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(58) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbWhite And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(59) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbWhite And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(60) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbWhite And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbGreen Then
Text1(61) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbWhite 
And Text1(63).BackColor = vbGreen Then
Text1(62) = a
End If
If Text1(55).BackColor = vbGreen And Text1(56).BackColor = vbGreen And Text1(57).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(59).BackColor = vbGreen And 
Text1(60).BackColor = vbGreen And Text1(61).BackColor = vbGreen And Text1(62).BackColor = vbGreen 
And Text1(63).BackColor = vbWhite Then
Text1(63) = a
End If
'radif hashtom
If Text1(64).BackColor = vbWhite And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(64) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbWhite And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(65) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbWhite And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(66) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbWhite And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(67) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbWhite And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(68) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbWhite And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(69) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbWhite And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbGreen Then
Text1(70) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbWhite 
And Text1(72).BackColor = vbGreen Then
Text1(71) = a
End If
If Text1(64).BackColor = vbGreen And Text1(65).BackColor = vbGreen And Text1(66).BackColor = 
vbGreen And Text1(67).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(69).BackColor = vbGreen And Text1(70).BackColor = vbGreen And Text1(71).BackColor = vbGreen 
And Text1(72).BackColor = vbWhite Then
Text1(72) = a
End If
'radif nohom
If Text1(73).BackColor = vbWhite And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(73) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbWhite And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(74) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbWhite And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(75) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbWhite And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(76) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbWhite And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(77) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbWhite And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(78) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbWhite And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbGreen Then
Text1(79) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbWhite 
And Text1(81).BackColor = vbGreen Then
Text1(80) = a
End If
If Text1(73).BackColor = vbGreen And Text1(74).BackColor = vbGreen And Text1(75).BackColor = 
vbGreen And Text1(76).BackColor = vbGreen And Text1(77).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen And Text1(79).BackColor = vbGreen And Text1(80).BackColor = vbGreen 
And Text1(81).BackColor = vbWhite Then
Text1(81) = a
End If
'baray moshakhas kardan adadha
For e = 1 To 81
If Text1(e) = "" Then
Text1(e).BackColor = vbWhite
End If
Next e
Next a
Timer5.Interval = 0
If Option1.Value = False Then
Timer6.Interval = 1
End If
End Sub
Private Sub Timer6_Timer()
For e = 1 To 81
If Text1(e) = "" Then
Text1(e).BackColor = vbWhite
End If
Next
For a = 1 To 9
For b = 1 To 81
c = 0
If Text1(b) = a Then
If b >= 1 And b <= 9 Then
 For d = 1 To 9
 Text1(d).BackColor = vbGreen
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 10 And b <= 18 Then
 c = -18
 For d = 10 To 18
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
 
End If
If b >= 19 And b <= 27 Then
 c = -27
 For d = 19 To 27
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
 
End If
If b >= 28 And b <= 36 Then
 c = -36
 For d = 28 To 36
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 37 And b <= 45 Then
 c = -45
 For d = 37 To 45
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 46 And b <= 54 Then
 c = -54
 For d = 46 To 54
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 55 And b <= 63 Then
 c = -63
 For d = 55 To 63
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 64 And b <= 72 Then
 c = -72
 For d = 64 To 72
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
If b >= 73 And b <= 81 Then
 c = -81
 For d = 73 To 81
 Text1(d).BackColor = vbGreen
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).BackColor = vbGreen
 End If
 
 Next d
End If
'bala samt chap
If (b >= 1 And b <= 3) Or (b >= 10 And b <= 12) Or (b >= 19 And b <= 21) Then
Text1(1).BackColor = vbGreen
Text1(2).BackColor = vbGreen
Text1(3).BackColor = vbGreen
Text1(10).BackColor = vbGreen
Text1(11).BackColor = vbGreen
Text1(12).BackColor = vbGreen
Text1(19).BackColor = vbGreen
Text1(20).BackColor = vbGreen
Text1(21).BackColor = vbGreen
 
End If
'bala
If (b >= 4 And b <= 6) Or (b >= 13 And b <= 15) Or (b >= 22 And b <= 24) Then
Text1(4).BackColor = vbGreen
Text1(5).BackColor = vbGreen
Text1(6).BackColor = vbGreen
Text1(13).BackColor = vbGreen
Text1(14).BackColor = vbGreen
Text1(15).BackColor = vbGreen
Text1(22).BackColor = vbGreen
Text1(23).BackColor = vbGreen
Text1(24).BackColor = vbGreen
End If
'bala samt rast
If (b >= 7 And b <= 9) Or (b >= 16 And b <= 18) Or (b >= 25 And b <= 27) Then
Text1(7).BackColor = vbGreen
Text1(8).BackColor = vbGreen
Text1(9).BackColor = vbGreen
Text1(16).BackColor = vbGreen
Text1(17).BackColor = vbGreen
Text1(18).BackColor = vbGreen
Text1(25).BackColor = vbGreen
Text1(26).BackColor = vbGreen
Text1(27).BackColor = vbGreen
End If
'chap
If (b >= 28 And b <= 30) Or (b >= 37 And b <= 39) Or (b >= 46 And b <= 48) Then
Text1(28).BackColor = vbGreen
Text1(29).BackColor = vbGreen
Text1(30).BackColor = vbGreen
Text1(37).BackColor = vbGreen
Text1(38).BackColor = vbGreen
Text1(39).BackColor = vbGreen
Text1(46).BackColor = vbGreen
Text1(47).BackColor = vbGreen
Text1(48).BackColor = vbGreen
End If
'vasat
If (b >= 31 And b <= 33) Or (b >= 40 And b <= 42) Or (b >= 49 And b <= 51) Then
Text1(31).BackColor = vbGreen
Text1(32).BackColor = vbGreen
Text1(33).BackColor = vbGreen
Text1(40).BackColor = vbGreen
Text1(41).BackColor = vbGreen
Text1(42).BackColor = vbGreen
Text1(49).BackColor = vbGreen
Text1(50).BackColor = vbGreen
Text1(51).BackColor = vbGreen
End If
'rast
If (b >= 34 And b <= 36) Or (b >= 43 And b <= 45) Or (b >= 52 And b <= 54) Then
Text1(34).BackColor = vbGreen
Text1(35).BackColor = vbGreen
Text1(36).BackColor = vbGreen
Text1(43).BackColor = vbGreen
Text1(44).BackColor = vbGreen
Text1(45).BackColor = vbGreen
Text1(52).BackColor = vbGreen
Text1(53).BackColor = vbGreen
Text1(54).BackColor = vbGreen
End If
'paean samt chap
If (b >= 55 And b <= 57) Or (b >= 64 And b <= 66) Or (b >= 73 And b <= 75) Then
Text1(55).BackColor = vbGreen
Text1(56).BackColor = vbGreen
Text1(57).BackColor = vbGreen
Text1(64).BackColor = vbGreen
Text1(65).BackColor = vbGreen
Text1(66).BackColor = vbGreen
Text1(73).BackColor = vbGreen
Text1(74).BackColor = vbGreen
Text1(75).BackColor = vbGreen
End If
'paean
If (b >= 58 And b <= 60) Or (b >= 67 And b <= 69) Or (b >= 76 And b <= 78) Then
Text1(58).BackColor = vbGreen
Text1(59).BackColor = vbGreen
Text1(60).BackColor = vbGreen
Text1(67).BackColor = vbGreen
Text1(68).BackColor = vbGreen
Text1(69).BackColor = vbGreen
Text1(76).BackColor = vbGreen
Text1(77).BackColor = vbGreen
Text1(78).BackColor = vbGreen
End If
'paean samt rast
If (b >= 61 And b <= 63) Or (b >= 70 And b <= 72) Or (b >= 79 And b <= 81) Then
Text1(61).BackColor = vbGreen
Text1(62).BackColor = vbGreen
Text1(63).BackColor = vbGreen
Text1(70).BackColor = vbGreen
Text1(71).BackColor = vbGreen
Text1(72).BackColor = vbGreen
Text1(79).BackColor = vbGreen
Text1(80).BackColor = vbGreen
Text1(81).BackColor = vbGreen
End If
End If
Next b
'baraye moghayesehe setoon
'setoon avval
If Text1(1).BackColor = vbWhite And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(1) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbWhite And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(10) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbWhite 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(19) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbWhite And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(28) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbWhite And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(37) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbWhite And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(46) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbWhite And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbGreen Then
Text1(55) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbWhite And 
Text1(73).BackColor = vbGreen Then
Text1(64) = a
End If
If Text1(1).BackColor = vbGreen And Text1(10).BackColor = vbGreen And Text1(19).BackColor = vbGreen 
And Text1(28).BackColor = vbGreen And Text1(37).BackColor = vbGreen And Text1(46).BackColor = 
vbGreen And Text1(55).BackColor = vbGreen And Text1(64).BackColor = vbGreen And 
Text1(73).BackColor = vbWhite Then
Text1(73) = a
End If
'setoon dovvom
If Text1(2).BackColor = vbWhite And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(2) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbWhite And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And
Text1(74).BackColor = vbGreen Then
Text1(11) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbWhite 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(20) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbWhite And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(29) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbWhite And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(38) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbWhite And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(47) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbWhite And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbGreen Then
Text1(56) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbWhite And 
Text1(74).BackColor = vbGreen Then
Text1(65) = a
End If
If Text1(2).BackColor = vbGreen And Text1(11).BackColor = vbGreen And Text1(20).BackColor = vbGreen 
And Text1(29).BackColor = vbGreen And Text1(38).BackColor = vbGreen And Text1(47).BackColor = 
vbGreen And Text1(56).BackColor = vbGreen And Text1(65).BackColor = vbGreen And 
Text1(74).BackColor = vbWhite Then
Text1(74) = a
End If
'setoon sevvom
If Text1(3).BackColor = vbWhite And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(3) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbWhite And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(12) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbWhite 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(21) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbWhite 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(21) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbWhite And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(30) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbWhite And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(39) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbWhite And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(48) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbWhite And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbGreen Then
Text1(57) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbWhite And 
Text1(75).BackColor = vbGreen Then
Text1(66) = a
End If
If Text1(3).BackColor = vbGreen And Text1(12).BackColor = vbGreen And Text1(21).BackColor = vbGreen 
And Text1(30).BackColor = vbGreen And Text1(39).BackColor = vbGreen And Text1(48).BackColor = 
vbGreen And Text1(57).BackColor = vbGreen And Text1(66).BackColor = vbGreen And 
Text1(75).BackColor = vbWhite Then
Text1(75) = a
End If
'setton charom
If Text1(4).BackColor = vbWhite And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(4) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbWhite And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(13) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbWhite 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(22) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbWhite And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(31) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbWhite And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(40) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbWhite And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(49) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbWhite And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbGreen Then
Text1(58) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbWhite And 
Text1(76).BackColor = vbGreen Then
Text1(67) = a
End If
If Text1(4).BackColor = vbGreen And Text1(13).BackColor = vbGreen And Text1(22).BackColor = vbGreen 
And Text1(31).BackColor = vbGreen And Text1(40).BackColor = vbGreen And Text1(49).BackColor = 
vbGreen And Text1(58).BackColor = vbGreen And Text1(67).BackColor = vbGreen And 
Text1(76).BackColor = vbWhite Then
Text1(76) = a
End If
'setoon panjom
If Text1(5).BackColor = vbWhite And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(5) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbWhite And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(14) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbWhite 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And
Text1(77).BackColor = vbGreen Then
Text1(23) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbWhite And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(32) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbWhite And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(41) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbWhite And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(50) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbWhite And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbGreen Then
Text1(59) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbWhite And 
Text1(77).BackColor = vbGreen Then
Text1(68) = a
End If
If Text1(5).BackColor = vbGreen And Text1(14).BackColor = vbGreen And Text1(23).BackColor = vbGreen 
And Text1(32).BackColor = vbGreen And Text1(41).BackColor = vbGreen And Text1(50).BackColor = 
vbGreen And Text1(59).BackColor = vbGreen And Text1(68).BackColor = vbGreen And 
Text1(77).BackColor = vbWhite Then
Text1(77) = a
End If
'setoon sheshom
If Text1(6).BackColor = vbWhite And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(6) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbWhite And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(15) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbWhite 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(24) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbWhite And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(33) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbWhite And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(42) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbWhite And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(51) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbWhite And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbGreen Then
Text1(60) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbWhite And 
Text1(78).BackColor = vbGreen Then
Text1(69) = a
End If
If Text1(6).BackColor = vbGreen And Text1(15).BackColor = vbGreen And Text1(24).BackColor = vbGreen 
And Text1(33).BackColor = vbGreen And Text1(42).BackColor = vbGreen And Text1(51).BackColor = 
vbGreen And Text1(60).BackColor = vbGreen And Text1(69).BackColor = vbGreen And 
Text1(78).BackColor = vbWhite Then
Text1(78) = a
End If
'setoon haftom
If Text1(7).BackColor = vbWhite And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(7) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbWhite And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(16) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbWhite 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(25) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbWhite And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(34) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbWhite And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(43) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbWhite And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(52) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbWhite And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbGreen Then
Text1(61) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbWhite And 
Text1(79).BackColor = vbGreen Then
Text1(70) = a
End If
If Text1(7).BackColor = vbGreen And Text1(16).BackColor = vbGreen And Text1(25).BackColor = vbGreen 
And Text1(34).BackColor = vbGreen And Text1(43).BackColor = vbGreen And Text1(52).BackColor = 
vbGreen And Text1(61).BackColor = vbGreen And Text1(70).BackColor = vbGreen And 
Text1(79).BackColor = vbWhite Then
Text1(79) = a
End If
'setoon hashtom
If Text1(8).BackColor = vbWhite And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(8) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbWhite And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(17) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbWhite 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(26) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbWhite And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(35) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbWhite And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(44) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbWhite And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(53) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbWhite And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbGreen Then
Text1(62) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbWhite And 
Text1(80).BackColor = vbGreen Then
Text1(71) = a
End If
If Text1(8).BackColor = vbGreen And Text1(17).BackColor = vbGreen And Text1(26).BackColor = vbGreen 
And Text1(35).BackColor = vbGreen And Text1(44).BackColor = vbGreen And Text1(53).BackColor = 
vbGreen And Text1(62).BackColor = vbGreen And Text1(71).BackColor = vbGreen And 
Text1(80).BackColor = vbWhite Then
Text1(80) = a
End If
'setoon nohom
If Text1(9).BackColor = vbWhite And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(9) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbWhite And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(18) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbWhite 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(27) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbWhite And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(36) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbWhite And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(45) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbWhite And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(54) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbWhite And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbGreen Then
Text1(63) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbWhite And 
Text1(81).BackColor = vbGreen Then
Text1(72) = a
End If
If Text1(9).BackColor = vbGreen And Text1(18).BackColor = vbGreen And Text1(27).BackColor = vbGreen 
And Text1(36).BackColor = vbGreen And Text1(45).BackColor = vbGreen And Text1(54).BackColor = 
vbGreen And Text1(63).BackColor = vbGreen And Text1(72).BackColor = vbGreen And 
Text1(81).BackColor = vbWhite Then
Text1(81) = a
End If
'baray moshakhas kardan adadha
For e = 1 To 81
If Text1(e) = "" Then
Text1(e).BackColor = vbWhite
End If
Next e
Next a
Timer6.Interval = 0
If Option1.Value = False Then
Timer7.Interval = 1
End If
End Sub
Private Sub Timer7_Timer()
For b = 1 To 81
For a = 1 To 81
Text1(a).DataMember = 0
Next a
For d = 1 To 9
Label1(d).Visible = True
Next d
c = 0
If b >= 1 And b <= 9 Then
 For d = 1 To 9
 Text1(d).DataMember = 1
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
If b >= 10 And b <= 18 Then
 c = -18
 For d = 10 To 18
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
 
End If
If b >= 19 And b <= 27 Then
 c = -27
 For d = 19 To 27
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
 
End If
If b >= 28 And b <= 36 Then
 c = -36
 For d = 28 To 36
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
If b >= 37 And b <= 45 Then
 c = -45
 For d = 37 To 45
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
If b >= 46 And b <= 54 Then
 c = -54
 For d = 46 To 54
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
If b >= 55 And b <= 63 Then
 c = -63
 For d = 55 To 63
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
If b >= 64 And b <= 72 Then
 c = -72
 For d = 64 To 72
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
If b >= 73 And b <= 81 Then
 c = -81
 For d = 73 To 81
 Text1(d).DataMember = 1
 
 c = c + 9
 
 If c <= 72 Then
 Text1(b + c).DataMember = 1
 End If
 
 Next d
End If
'bala samt chap
If (b >= 1 And b <= 3) Or (b >= 10 And b <= 12) Or (b >= 19 And b <= 21) Then
Text1(1).DataMember = 1
Text1(2).DataMember = 1
Text1(3).DataMember = 1
Text1(10).DataMember = 1
Text1(11).DataMember = 1
Text1(12).DataMember = 1
Text1(19).DataMember = 1
Text1(20).DataMember = 1
Text1(21).DataMember = 1
 
End If
'bala
If (b >= 4 And b <= 6) Or (b >= 13 And b <= 15) Or (b >= 22 And b <= 24) Then
Text1(4).DataMember = 1
Text1(5).DataMember = 1
Text1(6).DataMember = 1
Text1(13).DataMember = 1
Text1(14).DataMember = 1
Text1(15).DataMember = 1
Text1(22).DataMember = 1
Text1(23).DataMember = 1
Text1(24).DataMember = 1
End If
'bala samt rast
If (b >= 7 And b <= 9) Or (b >= 16 And b <= 18) Or (b >= 25 And b <= 27) Then
Text1(7).DataMember = 1
Text1(8).DataMember = 1
Text1(9).DataMember = 1
Text1(16).DataMember = 1
Text1(17).DataMember = 1
Text1(18).DataMember = 1
Text1(25).DataMember = 1
Text1(26).DataMember = 1
Text1(27).DataMember = 1
End If
'chap
If (b >= 28 And b <= 30) Or (b >= 37 And b <= 39) Or (b >= 46 And b <= 48) Then
Text1(28).DataMember = 1
Text1(29).DataMember = 1
Text1(30).DataMember = 1
Text1(37).DataMember = 1
Text1(38).DataMember = 1
Text1(39).DataMember = 1
Text1(46).DataMember = 1
Text1(47).DataMember = 1
Text1(48).DataMember = 1
End If
'vasat
If (b >= 31 And b <= 33) Or (b >= 40 And b <= 42) Or (b >= 49 And b <= 51) Then
Text1(31).DataMember = 1
Text1(32).DataMember = 1
Text1(33).DataMember = 1
Text1(40).DataMember = 1
Text1(41).DataMember = 1
Text1(42).DataMember = 1
Text1(49).DataMember = 1
Text1(50).DataMember = 1
Text1(51).DataMember = 1
End If
'rast
If (b >= 34 And b <= 36) Or (b >= 43 And b <= 45) Or (b >= 52 And b <= 54) Then
Text1(34).DataMember = 1
Text1(35).DataMember = 1
Text1(36).DataMember = 1
Text1(43).DataMember = 1
Text1(44).DataMember = 1
Text1(45).DataMember = 1
Text1(52).DataMember = 1
Text1(53).DataMember = 1
Text1(54).DataMember = 1
End If
'paean samt chap
If (b >= 55 And b <= 57) Or (b >= 64 And b <= 66) Or (b >= 73 And b <= 75) Then
Text1(55).DataMember = 1
Text1(56).DataMember = 1
Text1(57).DataMember = 1
Text1(64).DataMember = 1
Text1(65).DataMember = 1
Text1(66).DataMember = 1
Text1(73).DataMember = 1
Text1(74).DataMember = 1
Text1(75).DataMember = 1
End If
'paean
If (b >= 58 And b <= 60) Or (b >= 67 And b <= 69) Or (b >= 76 And b <= 78) Then
Text1(58).DataMember = 1
Text1(59).DataMember = 1
Text1(60).DataMember = 1
Text1(67).DataMember = 1
Text1(68).DataMember = 1
Text1(69).DataMember = 1
Text1(76).DataMember = 1
Text1(77).DataMember = 1
Text1(78).DataMember = 1
End If
'paean samt rast
If (b >= 61 And b <= 63) Or (b >= 70 And b <= 72) Or (b >= 79 And b <= 81) Then
Text1(61).DataMember = 1
Text1(62).DataMember = 1
Text1(63).DataMember = 1
Text1(70).DataMember = 1
Text1(71).DataMember = 1
Text1(72).DataMember = 1
Text1(79).DataMember = 1
Text1(80).DataMember = 1
Text1(81).DataMember = 1
End If
For a = 1 To 9
For b2 = 1 To 81
If Text1(b2).DataMember = 1 Then
If Text1(b2) = a Then
Label1(a).Visible = False
End If
End If
Next b2
Next a
c = 0
For d = 1 To 9
If Label1(d).Visible = False Then
c = c + 1
End If
Next d
If c = 8 And Text1(b) = "" Then
For d = 1 To 9
If Label1(d).Visible = True Then
Text1(b) = d
End If
Next d
End If
If c = 7 And Text1(b) = "" Then
For d = 1 To 9
If Label1(d).Visible = True Then
Text1(b).BackColor = vbYellow
End If
Next d
End If
If c = 6 And Text1(b) = "" Then
For d = 1 To 9
If Label1(d).Visible = True Then
Text1(b).BackColor = vbBlue
End If
Next d
End If
Next b
Timer7.Interval = 0
If Option1.Value = False Then
Timer0.Interval = 1
End If
End Sub
