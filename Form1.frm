VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Encryption / Decryption Program by : Paw Jershauge"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "EnCrypted text"
      Enabled         =   0   'False
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   9615
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   8
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Decrypt"
         Height          =   255
         Left            =   8160
         TabIndex        =   5
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   9375
      End
      Begin VB.Label Label2 
         Caption         =   "Decryption Key (Max 10) :"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   3000
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text to EnCrypt"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Encrypt"
         Height          =   255
         Left            =   8160
         TabIndex        =   2
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   9375
      End
      Begin VB.Label Label1 
         Caption         =   "Encryption Key (Max 10) :"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   2880
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'
'             Encryption program by Paw Jershauge (11/9-2000)
'
'                      Email : Viper_one@get2net.dk
'
'                   WebSite : Not up yet (Comming soon)
'
'                   Want HELP, just email me, OKAY.
'************************************************************************************************

Private Sub Command1_Click()

If Text3.Text = "" Then
    MsgBox "Must have a key, Just enter one", vbCritical, "Need some a key" 'If no key the ask for it
Else
    Frame1.Enabled = False 'Lock Encrypt frame from user
    Frame2.Enabled = True 'Unlock Decrypt frame from user
    Text2.Text = ""
    Encrypt Text1.Text, Text3.Text, Text2 'Encrypt function
End If

End Sub

Private Sub Command2_Click()
Dim Suc As Boolean
If Text4.Text = "" Then
    MsgBox "Must have a key, Just enter one", vbCritical, "Need some a key" 'If no key the ask for it
Else
    Frame2.Enabled = False 'Lock Decrypt frame from user
    Frame1.Enabled = True 'Unlock Encrypt frame from user
    Text1.Text = ""
    Suc = Decrypt(Text2.Text, Text4.Text, Text1)  ' Decrypt function
    If Suc = False Then
    Frame1.Enabled = False 'Lock Encrypt frame from user
    Frame2.Enabled = True 'Unlock Decrypt frame from user
    End If
End If

End Sub

Function Encrypt(EnCptText As String, Key As String, TxtBox As TextBox)
Dim Keylen As Integer, CharCode As Integer, TotalChar As Long

'*****************Calculate key**************
For K = 1 To Len(Key)
Keylen = Keylen + (Asc(Mid(Key, K, 1)) + 4)
Next K
'*****************End Calculating************

'*****************EnCrypt********************
For X = 1 To Len(EnCptText)
CharCode = (Asc(Mid(EnCptText, X, 1)) + Keylen)
TotalChar = TotalChar + CharCode
TxtBox.Text = TxtBox.Text & "." & CharCode 'Convert String to code
Next X

TxtBox.Text = TxtBox.Text & ".#" & TotalChar  'Add a seperator and a check number to the end of string.
'*****************End Encrypt****************
End Function

Function Decrypt(DeCptText As String, Key As String, TxtBox As TextBox) As Boolean
Dim Keylen As Integer, X, x1, x2
Dim CheckNumber As Long


'*****************Calculate key**************
For K = 1 To Len(Key)
Keylen = Keylen + (Asc(Mid(Key, K, 1)) + 4)
Next K
'*****************End Calculating************

'*****************DeCrypt********************
Do While X < Len(DeCptText) ' do loop to end of string
X = X + 1
    If "." = Mid(DeCptText, X, 1) Then 'if first seperator is found then
        x1 = X + 1
        Do While "." <> Mid(DeCptText, x1, 1) 'Do loop ontill next seperator
        x1 = x1 + 1
            If "." = Mid(DeCptText, x1, 1) Then ' When next seperator is found then
            x2 = (x1 - 1) - X
            If (Mid(DeCptText, X + 1, x2) - Keylen) > 255 Then 'If the decoded number is bigger then 255 then report failure
            MsgBox "The data failed to be Decodeds, Someone may have edit the data", vbCritical, "Failed"
            Text1.Text = "Failed"
            Decrypt = False
            Exit Function
            End If
            If (Mid(DeCptText, X + 1, x2) - Keylen) <= 0 Then 'If the decoded number is under 0 then report failure
            MsgBox "The data failed to be Decodeds, Someone may have edit the data", vbCritical, "Failed"
            Text1.Text = "Failed"
            Decrypt = False
            Exit Function
            End If
            TxtBox.Text = TxtBox.Text & Chr(Mid(DeCptText, X + 1, x2) - Keylen) 'Get Char code from string and make text
            CheckNumber = CheckNumber + Mid(DeCptText, X + 1, x2)
            X = x1 - 1
            End If
            DoEvents
            If x1 > Len(DeCptText) Then 'If only checknumber is left then
            If CheckNumber = GetChkSum(DeCptText) Then 'Compere the checknumbers
            MsgBox "The data was Decoded with success", vbInformation, "Success" 'if checknumber is OK then show Success
            Decrypt = True
            Else
            MsgBox "The data failed to be Decodeds", vbCritical, "Failed" 'if checknumber is NOT OK then show Failure
            Text1.Text = ""
            Decrypt = False
            End If
            Exit Function
            End If
        Loop
    End If
DoEvents
Loop
'***************End Decrypt*******************
End Function

Function GetChkSum(Cpttext As String) As Long
Do While q < Len(Cpttext)
q = q + 1
If "#" = Mid(Cpttext, (Len(Cpttext) - q), 1) Then
GetChkSum = Mid(Cpttext, ((Len(Cpttext) - q) + 1), q)
Exit Function
End If
DoEvents
Loop
End Function

Private Sub Form_Load()
Text1.Text = "Hallo there, Thanks for trying my Encryption / Decryption program. Hope you like it."
End Sub

'NOTE : This sub routine may be deleted, Its just so the keys are the same.
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
Text4.Text = Text3.Text
End Sub
'--------------------------------------------------------------------------

