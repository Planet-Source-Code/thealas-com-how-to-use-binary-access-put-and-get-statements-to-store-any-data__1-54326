VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Form1"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Write to ""test.dat"""
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read from ""test.dat"""
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' A file demo, (C) me, alas@eunet.yu
' how to manage memory in vb...

' A SHORT TUTOR:
' This is what you will learn:
' TO WRITE FILES:

'    Open App.Path & "\test.dat" For Binary Access Write As #1
'        Put #1, 1, Data
'    Close #1

' TO READ FILES;

'    Open App.Path & "\test.dat" For Binary Access Read As #1
'        Get #1, 1, Data
'    Close #1

' That's all, now have a look at the example

' You can use ReadFile and WriteFile apis, but in VB you dont need
' this at all


Option Explicit

' Here is the data
Private Type HDATA
    SubName As String
    DateOfBirth As Date
    Age As Integer ':)
End Type: Dim Data(1 To 4) As HDATA

Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long


Private Sub Command1_Click()
    ' We dont need to allocate anything, vb will just read the stuffs:
    Open App.Path & "\test.dat" For Binary Access Read As #1
        Get #1, 1, Data ' GET THE DATA FROM THE FILE
    Close #1
    
    ' Now we can use the structure
    Dim i%
    Debug.Print "==================TEST.DAT======================"
    For i = 1 To 3
        Debug.Print "NAME: " & Data(i).SubName
        Debug.Print "BIRTH YEAR: " & Year(Data(i).DateOfBirth)
        Debug.Print "AGE: " & Data(i).Age
        Debug.Print "========================================"
    Next i
End Sub

Private Sub Command2_Click()
    ' First allocate some memory
    Data(1).Age = 19
    Data(1).DateOfBirth = #6/14/1785#
    Data(1).SubName = "Blahoraya"
    
    Data(2).Age = 15
    Data(2).DateOfBirth = #6/14/1788#
    Data(2).SubName = "Spozzwxnt"
    
    Data(3).Age = 17
    Data(3).DateOfBirth = #6/14/1786#
    Data(3).SubName = "MecMeggDoode"
    
    ' We have written this to memory:
    ' "Blahoraya    €mäÀ      Spozzwxnt    €äãÀ  MecMeggDoode    à?äÀ"
    ' this is showed as ANSI string
    
    ' Now write the memo to a stream
    Open App.Path & "\test.dat" For Binary Access Write As #1
        Put #1, 1, Data ' SAVE THE DATA TO A FILE
    Close #1
    
    CalcMemo ' Show the messagebox for information about used bytes
    
End Sub

Private Sub CalcMemo()

    ' This example shows how to calculate how much memory your data is holding

    Dim i&, size&
    For i = 1 To 4
        ' increment the counter
        size = size + Len(Data(i).SubName) + Len(Data(i).DateOfBirth) + Len(Data(i).Age)
    Next i
    size = size + 4 ' add 4 bytes for size of the whole structure
    
    MsgBox "MEMORY USED:" & size & " bytes"

    ' the file size should match with this, but file will be 2-3 bytes bigger
    
End Sub

