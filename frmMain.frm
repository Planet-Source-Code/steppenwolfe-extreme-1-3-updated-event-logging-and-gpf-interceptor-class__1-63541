VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extreme! Logging and Exception Handler - by Steppenwolfe"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdControl 
      Caption         =   "Text View"
      Height          =   405
      Index           =   2
      Left            =   6900
      TabIndex        =   5
      Top             =   5880
      Width           =   1275
   End
   Begin RichTextLib.RichTextBox txtReport 
      Height          =   5355
      Left            =   90
      TabIndex        =   4
      Top             =   390
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   9446
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Raise Error"
      Height          =   405
      Index           =   3
      Left            =   8280
      TabIndex        =   3
      Top             =   5880
      Width           =   1275
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "HTML View"
      Height          =   405
      Index           =   1
      Left            =   5490
      TabIndex        =   2
      Top             =   5880
      Width           =   1275
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Application Log"
      Height          =   405
      Index           =   0
      Left            =   4110
      TabIndex        =   1
      Top             =   5880
      Width           =   1275
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, _
                                                   ByVal dwExceptionFlags As Long, _
                                                   ByVal nNumberOfArguments As Long, _
                                                   lpArguments As Long)

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private iRandNum    As Integer

'## Module ID 03

Public Sub RaiseAnException(lException As Long)

    RaiseException lException, 0, 0, 0

End Sub

Private Function InIDE() As Boolean
'//test for run state

Dim sClass      As String
Dim sBuffer     As String
Dim lPos        As Long

   sClass = String$(260, 0)
   lPos = GetClassName(hwnd, sClass, 260)
   lPos = InStr(sClass, vbNullChar)
   
   If (lPos > 0) Then
      sClass = Left$(sClass, lPos - 1)
   End If

   If InStr(sClass, "ThunderForm") = 1 Or InStr(sClass, "ThunderMDIForm") = 1 Then
      InIDE = True
   End If

End Function

Private Sub cmdControl_Click(Index As Integer)
'## 03/002

'//example of event/locale tracking
'//best to use a method that is
'//not obvious to user..
'//I use M/R/C format ex. 01/001/0001
'//first number is module id named in header of
'//each module. Next is routine number,
'//each tracked sub/function is named.
'//last is for inter call logging in areas
'//where exceptions are most likely.
'//it is really only necessary to track areas
'//where faults are likely to occur, ex
'//complex calculations/api, subclassing etc..
'//For example, this form would be Module 3 - or 03
'//this routine is named 002, and in this routine there
'//are four mappings 0001, 0002, 0003 and 0004
'//so error id would be 03/002/000x
'//this makes it very simple in tracking down
'//the specific source of an event using
'//the error log..

    Select Case Index
    '//send to textbox
    Case 0
'## 03/002/0001
With cLog
    .ELocale = "03/002/0001"
    .EData = "event processing data - check"
End With

    '//text log
    txtReport.Text = vbNullString
    cLog.Log_Text txtReport
    
    '//as web page
    Case 1
'## 03/002/002
With cLog
    .ELocale = "03/002/0002"
    .EData = "event processing data - check"
End With

    '//html report
    cLog.Report_Web Me

    '//as text file
    Case 2
'## 03/002/003
With cLog
    .ELocale = "03/002/0003"
    .EData = "event processing data - check"
End With

    '//notepad report
    cLog.Report_Note
    
    '//force a crash
    Case 3
'## 03/002/004
With cLog
    .ELocale = "03/002/0004"
    .EData = "event processing data - check"
End With

    '//ide/runtime crash
    If InIDE Then
        RaiseAnException 9
    Else
        Dummy_Routine
    End If

    End Select
    
End Sub

'//generate and track error through subroutines
'//If compiled use the routines below..
'//otherwise uses raise error api

Private Sub Dummy_Routine()

Dim i(0 To 1) As Long


'## 03/004/001
With cLog
    .ELocale = "03/004/0001"
    .EData = "event processing data - check"
End With

'//use the error property structure to track position
'//of a fault within a routine
'//example will result in divide by zero error
'//on compiled app

    With cLog
        .ELocale = "Dummy_Routine"
        .EData = "some info about routine"
    End With

    '//raise an error
    i(555555) = 1
    
    
    '//add a processing event to app log
    cLog.Log_Event "Dummy_Routine", "no err"

End Sub
