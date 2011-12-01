VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} searchResult 
   Caption         =   "UserForm1"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   OleObjectBlob   =   "searchResult.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "searchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function initialize( _
        title As String, author As String, publisher As String, _
        results As Variant)
    Me.textTitle.Value = title
    Me.textAuthor.Value = author
    Me.textPublisher.Value = publisher
    
    Dim i As Integer
    For i = 0 To 4
        With Me.Controls("OptionButton" & i)
            If (i <= UBound(results)) Then
                .Caption = results(i)("title") & vbLf _
                    & results(i)("author") & vbLf _
                    & Left(results(i)("creators"), 20) & vbLf _
                    & results(i)("publisher") & " " & results(i)("publicationDate") & " " & results(i)("binding") & vbLf _
                    & results(i)("ean")
                On Error GoTo 0
            Else
                .Visible = False
                .Enabled = False
            End If
        End With
    Next
End Function

Private Sub CommandButton1_Click()
    Dim attributes As Object
    Set attributes = CreateObject("Scripting.Dictionary")
    
    Dim strs() As String
    Dim obj As Variant
    For Each obj In Me.Frame2.Controls
        If (obj.Value) Then ' TODO ここから再開
            Me.Tag = Right(obj.Name, 1)
        End If
    Next
    Me.Hide
End Sub

Private Sub CommandButton2_click()
    Me.Tag = "cancel"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Tag = "cancel"
    Me.Hide
End Sub
