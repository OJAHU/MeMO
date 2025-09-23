VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "メモ"
   ClientHeight    =   4440
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7080
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim c_y As Long
    c_y = 3
    Do While Cells(c_y, 6) <> ""
        c_y = c_y + 1
    Loop
    
    '内容が書かれているなら
    If TextBox2.Text <> "" Then
        '締切日を入力
        If TextBox1.Text = "今日" Then
            Cells(c_y, 5).Value = Format(Date, "mm/dd")
            Cells(c_y, 6).Value = TextBox2.Text
        ElseIf TextBox1.Value = "明日" Then
            Cells(c_y, 5).Value = DateAdd("d", 1, Date)
        ElseIf TextBox1.Value = "来週" Then
            Cells(c_y, 5).Value = DateAdd("d", 7, Date)
        Else
            Cells(c_y, 5).Value = TextBox1.Text
        End If
        
        '内容を入力
        Cells(c_y, 6).Value = TextBox2.Text
        
        '各タスクの終了ボタン作成
        With ActiveSheet.Buttons.Add(Cells(c_y, 4).Left, _
                                    Cells(c_y, 4).Top, _
                                    Cells(c_y, 4).Width, _
                                    Cells(c_y, 4).Height)
                                        .name = c_y
                                        .OnAction = "fin_task"
                                        .Characters.Text = "終了"
        End With
        
        'メモ数を増やす
        Range("B9").Value = Range("B9").Value + 1
        
        ActiveWorkbook.Save
        Unload Me
    Else
        Dim ans As String
        ans = MsgBox("内容を入力してください", Buttons:=vbOKOnly, Title:="エラー")
    End If
End Sub

