VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} isbnInputForm 
   Caption         =   "Amazonデータ取得を自動実行"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   OleObjectBlob   =   "isbnInputForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "isbnInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub isbn_Change()
    If Len(isbn.Value) < 3 Then
        Exit Sub
    ElseIf (Left(isbn.Value, 3) = "978") And (Len(isbn.Value) < 13) Then
        Exit Sub
    ElseIf (Len(isbn.Value) < 10) Then
        Exit Sub
    End If
    
    Cells(ActiveCell.row, 1).Value = isbn.Value
    Application.Cursor = xlWait
    main.setBookInfo
    Cells(ActiveCell.row + 1, 1).Select
    isbn.Value = ""
    Application.Cursor = xlDefault
    isbn.SetFocus
    
End Sub

Private Sub UserForm_Initialize()
    isbn.SetFocus
End Sub
