VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CD Reminder beta 1.1"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2760
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "CD Information"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Text            =   "Yes / No"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Text            =   "Warez"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Year:"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Month:"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Day:"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Date Borrowed:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Casing:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_save 
         Caption         =   "S&ave"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_about 
      Caption         =   "&About"
      Begin VB.Menu mnu_help 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnu_me 
         Caption         =   "A&bout Me"
      End
   End
   Begin VB.Menu mnu_search 
      Caption         =   "&Search"
      Begin VB.Menu mnu_name 
         Caption         =   "by &Name"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
iniPath$ = App.Path & "/cd.ini"
Pull$ = Text2.Text

entry$ = Text2.Text
r% = WritePrivateProfileString(Pull$, "Name", entry$, iniPath$)

entry$ = Text1.Text
r% = WritePrivateProfileString(Pull$, "Title", entry$, iniPath$)

entry$ = Combo1.Text
r% = WritePrivateProfileString(Pull$, "Type", entry$, iniPath$)

entry$ = Combo3.Text & " / " & Combo4.Text & " / " & Combo5.Text
r% = WritePrivateProfileString(Pull$, "Date", entry$, iniPath$)

entry$ = Combo2.Text
r% = WritePrivateProfileString(Pull$, "Casing", entry$, iniPath$)

Text3.Text = "New Entry Added"
End Sub

Private Sub Form_Load()
Combo1.AddItem "Music"
Combo1.AddItem "Game"
Combo1.AddItem "Application"
Combo1.AddItem "Warez"
Combo2.AddItem "Yes"
Combo2.AddItem "No"
Combo3.AddItem ("01")
Combo3.AddItem ("02")
Combo3.AddItem ("03")
Combo3.AddItem ("04")
Combo3.AddItem ("05")
Combo3.AddItem ("06")
Combo3.AddItem ("07")
Combo3.AddItem ("08")
Combo3.AddItem ("09")
Combo3.AddItem ("10")
Combo3.AddItem ("11")
Combo3.AddItem ("12")
Combo3.AddItem ("13")
Combo3.AddItem ("14")
Combo3.AddItem ("15")
Combo3.AddItem ("16")
Combo3.AddItem ("17")
Combo3.AddItem ("18")
Combo3.AddItem ("19")
Combo3.AddItem ("20")
Combo3.AddItem ("21")
Combo3.AddItem ("22")
Combo3.AddItem ("23")
Combo3.AddItem ("24")
Combo3.AddItem ("25")
Combo3.AddItem ("26")
Combo3.AddItem ("27")
Combo3.AddItem ("28")
Combo3.AddItem ("29")
Combo3.AddItem ("30")
Combo3.AddItem ("31")
Combo4.AddItem ("January")
Combo4.AddItem ("February")
Combo4.AddItem ("March")
Combo4.AddItem ("April")
Combo4.AddItem ("May")
Combo4.AddItem ("June")
Combo4.AddItem ("July")
Combo4.AddItem ("August")
Combo4.AddItem ("September")
Combo4.AddItem ("October")
Combo4.AddItem ("November")
Combo4.AddItem ("December")
Combo5.AddItem ("2000")
Combo5.AddItem ("2001")
Combo5.AddItem ("2002")
Combo5.AddItem ("2003")
Combo5.AddItem ("2004")
Combo5.AddItem ("2005")
Combo5.AddItem ("2006")
Combo5.AddItem ("2007")
Combo5.AddItem ("2008")
Combo5.AddItem ("2009")
Combo5.AddItem ("2010")
End Sub

Private Sub mnu_exit_Click()
x = MsgBox("Are you sure you want to quit? All unsaved information will be lost.", vbYesNo + vbInformation, "Are you sure you want to quit?")
If x = vbYes Then
End
Else:
End If
End Sub

Private Sub mnu_help_Click()
x = MsgBox("Simply fill in the information in all the applicable fields. To view information, open the text file of the person with the information which you want to view. All text files are located in this applications directory.", vbOKOnly + vbInformation, "Help")
End Sub

Private Sub mnu_me_Click()
x = MsgBox("This program was made by Casey in Visual Basic 6.0. I'd like to thank ^Funny^ A.K.A Pat, for all his help.", vbOKOnly + vbInformation, "About Casey")
End Sub

Private Sub mnu_name_Click()
Dim name, title, msg, typee, datee, casing As String
msg = InputBox("Enter the name of the person you wish to search for", "Search by Name")
iniPath$ = App.Path & "/CD.ini"
Pull$ = msg

name = GetFromINI(Pull$, "Name", iniPath$)
title = GetFromINI(Pull$, "Title", iniPath$)
typee = GetFromINI(Pull$, "Type", iniPath$)
casing = GetFromINI(Pull$, "Casing", iniPath$)
datee = GetFromINI(Pull$, "Date", iniPath$)


Text3.Text = "Name:" & Chr$(32) & name & vbCrLf & "Title:" & Chr$(32) & title & vbCrLf & "Type:" & Chr$(32) & typee & vbCrLf & "Casing:" & Chr$(32) & casing & vbCrLf & "Date:" & Chr$(32) & datee





End Sub

Private Sub mnu_print_Click()
Printer.FontSize = 12
Printer.Print "CD Reminding Program - Created by Casey, Idea by ^Funny^ A.K.A Pat"
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print "-------------------"
Printer.Print "Name: " & Text2.Text
Printer.Print "CD Title: " & Text1.Text
Printer.Print "CD Type: " & Combo1.Text
Printer.Print "Casing: " & Combo2.Text
Printer.Print "Date Borrowed: " & Combo3.Text & "/" & Combo4.Text & "/" & Combo5.Text
Printer.Print "-------------------"
Printer.EndDoc

End Sub

Private Sub mnu_save_Click()
iniPath$ = App.Path & "/cd.ini"
Pull$ = Text2.Text

entry$ = Text2.Text
r% = WritePrivateProfileString(Pull$, "Name", entry$, iniPath$)

entry$ = Text1.Text
r% = WritePrivateProfileString(Pull$, "Title", entry$, iniPath$)

entry$ = Combo1.Text
r% = WritePrivateProfileString(Pull$, "Type", entry$, iniPath$)

entry$ = Combo3.Text & " / " & Combo4.Text & " / " & Combo5.Text
r% = WritePrivateProfileString(Pull$, "Date", entry$, iniPath$)

entry$ = Combo2.Text
r% = WritePrivateProfileString(Pull$, "Casing", entry$, iniPath$)

End Sub


