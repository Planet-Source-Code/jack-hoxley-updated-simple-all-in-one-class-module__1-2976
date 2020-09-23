VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label lblcontents 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As New AllinOne

Private Sub Form_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
Dim temp As Integer
temp = f.GetHdiskSpace("C", True)
lblcontents = ">" & "Space on Drive C is " & temp & " Mb" & Chr$(13)
Dim Temp1 As Boolean
Temp1 = f.GetIDEmode
lblcontents = lblcontents & ">" & "Is this running in the IDE?  " & Temp1 & Chr$(13)

lblcontents = lblcontents & ">" & "*************************************" & Chr$(13)
temp = f.GetMemLoad
lblcontents = lblcontents & ">" & "Memory Load is at " & temp & "%" & Chr$(13)
Dim temp2 As Long
temp2 = f.GetPhysMemTotal
lblcontents = lblcontents & ">" & "Total Physical Memory = " & temp2 & " Bytes" & Chr$(13)
temp2 = f.GetPhysMemAvailable
lblcontents = lblcontents & ">" & "Physical Memory Available = " & temp2 & " Bytes" & Chr$(13)
temp2 = f.GetVirtMemTotal
lblcontents = lblcontents & ">" & "Total Virtual Memory = " & temp2 & " Bytes" & Chr$(13)
temp2 = f.GetVirtMemAvailable
lblcontents = lblcontents & ">" & "Virtual Memory Available = " & temp2 & " Bytes" & Chr$(13)
temp2 = f.GetPageFileMemTotal
lblcontents = lblcontents & ">" & "Total PageFile Memory = " & temp2 & " Bytes" & Chr$(13)
temp2 = f.GetPageFileMemAvailable
lblcontents = lblcontents & ">" & "PageFile Memory Available = " & temp2 & " Bytes" & Chr$(13)
lblcontents = lblcontents & ">" & "*************************************" & Chr$(13)
Dim temp3 As String
temp3 = f.GetWinComputerName
lblcontents = lblcontents & ">" & "Windows Computer Name is " & temp3 & Chr$(13)
temp3 = f.GetWinDisplayColors("CLS_BITS")
lblcontents = lblcontents & ">" & "Number of Bits on the Current Device: " & temp3 & Chr$(13)
temp3 = f.GetWinDisplayColors("CLS_COL")
lblcontents = lblcontents & ">" & "Number of Colours on the Current Device: " & temp3 & Chr$(13)
temp = f.GetWinResX
lblcontents = lblcontents & ">" & "Current Windows resolution (X): " & temp & Chr$(13)
temp = f.GetWinResY
lblcontents = lblcontents & ">" & "Current Windows resolution (Y): " & temp & Chr$(13)
Dim temp4 As String
temp4 = f.GetWinResXY
lblcontents = lblcontents & ">" & "Current Windows resolution (XY): " & temp4 & Chr$(13)
temp4 = f.GetWinVersion
lblcontents = lblcontents & ">" & "Windows Version: " & temp4 & Chr$(13)
lblcontents = lblcontents & ">" & "" & Chr$(13)
lblcontents = lblcontents & ">" & "*************************************" & Chr$(13)
lblcontents = lblcontents & ">" & "Please Send any bug reports to:" & Chr$(13)
lblcontents = lblcontents & ">" & "JollyJeffers@Hotmail.com" & Chr$(13)
lblcontents = lblcontents & ">" & "I know there are a lot of bugs but please bear with them" & Chr$(13)
lblcontents = lblcontents & ">" & "If you are intelligent you can see the potential of one massive reusable Class" & Chr$(13)
lblcontents = lblcontents & ">" & "Module" & Chr$(13)
lblcontents = lblcontents & ">" & "*************************************" & Chr$(13)
lblcontents = lblcontents & ">" & "Some Uses for this code" & Chr$(13)
lblcontents = lblcontents & ">" & "As you can see, this almost looks like a DOS prompt," & Chr$(13)
lblcontents = lblcontents & ">" & "I used this once in front of a game" & Chr$(13)
lblcontents = lblcontents & ">" & "And in front of a graphics package" & Chr$(13)
lblcontents = lblcontents & ">" & "It warned people if there werent enough colors or" & Chr$(13)
lblcontents = lblcontents & ">" & "High/low enough resolution" & Chr$(13)
lblcontents = lblcontents & ">" & "*************************************" & Chr$(13)
lblcontents = lblcontents & ">" & "Distribute this could, mangle it as much as you like," & Chr$(13)
lblcontents = lblcontents & ">" & "But..........." & Chr$(13)
lblcontents = lblcontents & ">" & "Send me an email, i like to know how many people are" & Chr$(13)
lblcontents = lblcontents & ">" & "Using my code" & Chr$(13)
lblcontents = lblcontents & ">" & "*************************************" & Chr$(13)
End Sub


Private Sub Form_Resize()
lblcontents.Top = 0
lblcontents.Left = 0
lblcontents.Width = Me.ScaleWidth
lblcontents.Height = Me.ScaleHeight
End Sub


Private Sub lblcontents_Click()
Unload Me
End
End Sub


