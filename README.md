# VBA-CfaR
Cash Flow at Risk VBA



Userform Code

Private Sub CommandButton1_Click()

Dim cfar As Worksheet
Set cfar = Sheets("CFAR")
Dim gbpusd As Worksheet
Set gbpusd = Sheets("GBPUSD")
Dim spot As Double
Let spot = gbpusd.Range("e2")

cfar.Range("B1").Value = ComboBox1
cfar.Range("B2").Value = ComboBox2
cfar.Range("C1").Value = TextBox1

Unload Me

If ComboBox1.Value = "Payable" Then
cfar.Range("D1").Formula = "=c1*c2"
Else
cfar.Range("d1").Formula = "=c1/c2"

End If

Call Macro1

Unload Me

End Sub


Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub UserForm_Initialize()

With UserForm1.ComboBox1

.AddItem "Payable"
.AddItem "Receivable"

End With

With UserForm1.ComboBox2

.AddItem "GBPUSD"
.AddItem "EURUSD"

End With



End Sub

Private Sub UserForm_Click()

End Sub
