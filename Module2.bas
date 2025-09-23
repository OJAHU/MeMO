Attribute VB_Name = "Module2"
Sub write_memo() 'タスク作成
    Load UserForm1 'ユーザーフォームの呼び出し
    UserForm1.Show
End Sub
Sub close_memo() 'メモを閉じる
    'メモ数が0なら
    If Range("B9").Value = 0 Then
        Dim name As String
        name = Range("A1").Value
            
        '削除ウインドウ
        Dim ans As Variant
        ans = ActiveSheet.Delete
        
        'ウインドウで削除を押したら
        If ans = True Then
            Dim c_y As Long
            c_y = 3
            
            'メモ名
        
            Dim y As Long
            
            Do While Worksheets("Info").Cells(c_y, 7).Value <> "" Or Worksheets("Info").Cells(c_y, 8).Value <> ""
                '教科のメモなら
                If name = Worksheets("Info").Cells(c_y, 7).Value Then
                    'Infoシートの対応するものを空白にする
                    Worksheets("Info").Cells(c_y, 7).Value = ""
                    
                    y = c_y + 1
                    
                    'Infoシートの空欄部分をなくすようにずらす
                    Do While Worksheets("Info").Cells(y, 7).Value <> ""
                        Worksheets("Info").Cells(y - 1, 7).Value = Worksheets("Info").Cells(y, 7).Value
                        Worksheets("Info").Cells(y, 7).Value = ""
                        
                        y = y + 1
                    Loop
                    
                    Exit Do
                    
                    
                'その他のメモなら
                ElseIf name = Worksheets("Info").Cells(c_y, 8).Value Then
                    'Infoシートの対応するものを空白にする
                    Worksheets("Info").Cells(c_y, 8).Value = ""
                    
                    y = c_y + 1
                    
                    'Infoシートの空欄部分をなくすようにずらす
                    Do While Worksheets("Info").Cells(y, 8).Value <> ""
                        Worksheets("Info").Cells(y - 1, 8).Value = Worksheets("Info").Cells(y, 8).Value
                        Worksheets("Info").Cells(y, 8).Value = ""
                        
                        y = y + 1
                    Loop
                    
                    Exit Do
                End If
                
                c_y = c_y + 1
            Loop
        End If
    Else
        Dim rslt As VbMsgBoxResult
        rslt = MsgBox("タスクをすべて終了してから閉じてください", Buttons:=vbCritical, Title:="エラー")
    End If
    
    ActiveWorkbook.Save
End Sub
Sub go_home() '時間割シートに戻る
    Worksheets("時間割").Activate
End Sub
Sub fin_task() 'メモ上のタスクを削除
    Dim rslt As VbMsgBoxResult
    rslt = MsgBox("タスクを終了しますか？", Buttons:=vbYesNo, Title:="終了")
    
    If rslt = vbYes Then
        Dim btn As Object
        Dim pushed As Long '押されたボタンのy座標を取得
        pushed = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row
        
        Range(Cells(pushed, 5), Cells(pushed, 6)).Value = ""
        Range(Cells(pushed, 5), Cells(pushed, 6)).Interior.ColorIndex = 0
           
        '終了ボタンを削除
        For Each btn In ActiveSheet.Buttons
            If IsNumeric(btn.name) Then
                btn.Delete
            End If
        Next btn
        
        Dim c_y As Long
        c_y = pushed + 1
        
        '空白部分をずらし埋める
        Do While Cells(c_y, 6) <> ""
            Range(Cells(c_y - 1, 5), Cells(c_y - 1, 6)).Value = Range(Cells(c_y, 5), Cells(c_y, 6)).Value
            Range(Cells(c_y, 5), Cells(c_y, 6)) = ""
            Range(Cells(c_y, 5), Cells(c_y, 6)).Interior.ColorIndex = 0
            
            c_y = c_y + 1
        Loop
        
        c_y = 3
        Do While Cells(c_y, 6) <> ""
            With ActiveSheet.Buttons.Add(Cells(c_y, 4).Left, _
                                            Cells(c_y, 4).Top, _
                                            Cells(c_y, 4).Width, _
                                            Cells(c_y, 4).Height)
                                                .name = c_y
                                                .OnAction = "fin_task"
                                                .Characters.Text = "終了"
            End With
            c_y = c_y + 1
        Loop
        
        Range("B9").Value = Range("B9").Value - 1
    End If
End Sub
Sub change() 'メモの移動
    Dim subject As String
    subject = ActiveCell
    
    '既にメモが作成されているならそれを開く
    On Error GoTo ErrLabel
    Worksheets(subject).Activate
    
    'メモ一覧の削除
    Application.DisplayAlerts = False
    Worksheets("メモ一覧").Delete
    Application.DisplayAlerts = True
    Exit Sub
    
ErrLabel:
    Dim ans As VbMsgBoxResult
    ans = MsgBox("アクティブセルを調整してください", Buttons:=vbCritical, Title:="エラー")
End Sub
Sub ws_delete()
    'メモ一覧の削除
    Application.DisplayAlerts = False
    Worksheets("メモ一覧").Delete
    Application.DisplayAlerts = True
End Sub


