Attribute VB_Name = "ボタン表示切替"
Option Explicit

Public Importance As String
Public Version As String
Sub 選択バージョン取得()
    '選択中のバージョンを取得、PROは「○」、MECHは「×」を返す
    With Sheets("進捗管理").Shapes("バージョンボタン").TextFrame2.TextRange.Characters
        If .Text = "MECH" Then
            Version = "×"
        ElseIf .Text = "IJCAD" Then
            Version = "○"
        End If
    End With
End Sub

Sub バージョンボタン表示()

    With ActiveSheet.Shapes("バージョンボタン").TextFrame2.TextRange.Characters
        If .Text = "MECH" Then
            .Text = "IJCAD"
        ElseIf .Text = "IJCAD" Then
            .Text = "MECH"
        End If
    End With
End Sub

Sub 選択重要度取得()
    '選択中の重要度を取得
    Importance = Sheets("進捗管理").Shapes("重要度ボタン").TextFrame2.TextRange.Characters.Text

End Sub

Sub 重要度ボタン表示()

    With ActiveSheet.Shapes("重要度ボタン").TextFrame2.TextRange.Characters
        If .Text = "A" Then
            .Text = "B"
        ElseIf .Text = "B" Then
            .Text = "C"
        ElseIf .Text = "C" Then
            .Text = "A"
        End If
    End With
    



End Sub


