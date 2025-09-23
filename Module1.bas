Attribute VB_Name = "Module1"
Sub 教科確定ボタン_Click()
    '教科未登録なら
    Dim ans As VbMsgBoxResult
    
    If Worksheets("Info").Cells(2, 2).Value = 0 Then
        ans = MsgBox("教科を確定しますか？", Buttons:=vbYesNo + vbQuestion, Title:="確認")
        
        If ans = vbYes Then
            Dim c_y As Long
            c_y = 3
            
            For x = 3 To 7
                For y = 3 To 8
                    '時間割表から教科と対応する座標をInfoシートに登録
                    If Worksheets("時間割").Cells(y, x).Value <> "" Then
                        Worksheets("Info").Cells(c_y, 2).Value = Worksheets("時間割").Cells(y, x).Value
                        Worksheets("Info").Cells(c_y, 4).Value = x
                        Worksheets("Info").Cells(c_y, 5).Value = y
                        
                        c_y = c_y + 1
                    End If
                Next
            Next
            
            ans = MsgBox("教科の登録が完了しました" & vbCrLf & "教室の登録を実施してください", Buttons:=vbInformation, Title:="登録")
            '教科を登録状態にする
            Worksheets("Info").Cells(2, 2).Value = 1
            
            '保存
            ActiveWorkbook.Save
        End If
    Else
        ans = MsgBox("教科は既に登録されています", Buttons:=vbCritical, Title:="登録済")
    End If
End Sub
Sub 教室確定ボタン_Click()
    Dim ans As VbMsgBoxResult
    
    '教科が登録状態で教室が未登録なら
    If Worksheets("Info").Cells(2, 3).Value = 0 And Worksheets("Info").Cells(2, 2).Value = 1 Then
        Dim c_y As Long
        c_y = 3
        
        Do While Worksheets("Info").Cells(c_y, 2).Value <> ""
            '対応する教科の時間割シートにおいてのxy座標
            Dim x As Long
            x = Worksheets("Info").Cells(c_y, 4).Value
            Dim y As Long
            y = Worksheets("Info").Cells(c_y, 5).Value
            '対応するxy座標の背景色を黄色にする
            Worksheets("時間割").Cells(y, x).Interior.ColorIndex = 6
            
            classroom = InputBox(Worksheets("Info").Cells(c_y, 2).Value & "の教室を入力してください")
            Worksheets("Info").Cells(c_y, 3).Value = classroom
            
            '背景色をもとに戻す
            Worksheets("時間割").Cells(y, x).Interior.ColorIndex = 0
                        
            c_y = c_y + 1
        Loop
        
        ans = MsgBox("教室の登録が完了しました", Buttons:=vbInformation, Title:="登録")
        '教室を登録状態にする
        Worksheets("Info").Cells(2, 3).Value = 1
        
        '保存
        ActiveWorkbook.Save
    '教室未登録なら
    ElseIf Worksheets("Info").Cells(2, 2).Value = 0 Then
        ans = MsgBox("先に教科の登録をしてください", Buttons:=vbCritical, Title:="エラー")
    Else
        ans = MsgBox("教室は既に登録されています", Buttons:=vbCritical, Title:="登録済")
    End If
End Sub
Sub 初期化_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("初期化しますか？", Buttons:=vbYesNo + vbExclamation, Title:="確認")
    
    If ans = vbYes Then
        '登録状態の初期化
        Worksheets("Info").Cells(2, 2).Value = 0
        Worksheets("Info").Cells(2, 3).Value = 0
        
        Dim c_y As Long
        c_y = 3
        
        For x = 3 To 7
            For y = 3 To 8
                '背景色の初期化
                Worksheets("時間割").Cells(y, x).Interior.ColorIndex = 0
                
                If Worksheets("時間割").Cells(y, x).Value <> "" Then
                    '時間割シートの教科を初期化
                    Worksheets("時間割").Cells(y, x).Value = ""
                    
                    'Infoシートの情報を初期化
                    Worksheets("Info").Cells(c_y, 2).Value = ""
                    Worksheets("Info").Cells(c_y, 3).Value = ""
                    Worksheets("Info").Cells(c_y, 4).Value = ""
                    Worksheets("Info").Cells(c_y, 5).Value = ""
                        
                    c_y = c_y + 1
                End If
            Next
        Next
        
        '保存
        ActiveWorkbook.Save
    End If
End Sub
Sub 教科切り替えボタン_Click()
    '登録状態なら
    If Worksheets("Info").Cells(2, 2).Value = 1 And Worksheets("Info").Cells(2, 3).Value = 1 Then
        Dim c_y As Long
        Dim x As Long
        Dim y As Long
        
        '現在が教科表示状態なら
        If Worksheets("Info").Cells(2, 6).Value = "sbj" Then
            c_y = 3
            
            Do While Worksheets("Info").Cells(c_y, 2).Value <> ""
                x = Worksheets("Info").Cells(c_y, 4).Value
                y = Worksheets("Info").Cells(c_y, 5).Value
                
                '時間割シートの対応する教科を教室にする
                Worksheets("時間割").Cells(y, x).Value = Worksheets("Info").Cells(c_y, 3).Value
                c_y = c_y + 1
            Loop
            '教室表示状態にする
            Worksheets("Info").Cells(2, 6).Value = "cls"
        '現在が教室表示状態なら
        ElseIf Worksheets("Info").Cells(2, 6).Value = "cls" Then
            c_y = 3
            
            Do While Worksheets("Info").Cells(c_y, 3).Value <> ""
                x = Worksheets("Info").Cells(c_y, 4).Value
                y = Worksheets("Info").Cells(c_y, 5).Value
                
                '時間割シートの対応する教室を教科にする
                Worksheets("時間割").Cells(y, x).Value = Worksheets("Info").Cells(c_y, 2).Value
                c_y = c_y + 1
            Loop
            '教科表示状態にする
            Worksheets("Info").Cells(2, 6).Value = "sbj"
        End If
    Else
        Dim ans As VbMsgBoxResult
        ans = MsgBox("教科・教室登録を完了してください", Buttons:=vbCritical, Title:="エラー")
    End If
End Sub
Sub メモ作成ボタン_Click() '教科用メモ作成
    'アクティブセルを読み込む
    Dim target As String
    target = ActiveCell.Value
    
    Dim c_y As Long
    c_y = 3
    
    '教室が指定されているなら教科に変換する
    Do While Worksheets("Info").Cells(c_y, 3).Value <> ""
        If target = Worksheets("Info").Cells(c_y, 3).Value Then
            target = Worksheets("Info").Cells(c_y, 2).Value
            Exit Do
        End If
        
        c_y = c_y + 1
    Loop
    
    'メモが既に作成されているならそれをアクティブにし、処理を終了する
    For Each memo In Worksheets
        If memo.name = target Then
            memo.Activate
            Exit Sub
        End If
    Next memo
    
    c_y = 3
    
    'Infoシートに登録された教科と一致するまで
    Do While Worksheets("Info").Cells(c_y, 2).Value <> ""
        'アクティブセルが教科と教室のどちらかが一致するなら
        If Worksheets("Info").Cells(c_y, 2).Value = target Or Worksheets("Info").Cells(c_y, 3).Value = target Then
            Dim ws As Worksheet
            Set ws = Sheets.Add(After:=Sheets(Sheets.Count)) 'シートを最後尾に追加
            
            'シート名
            ws.name = target
            'タスク入力ボタン作成
            With ActiveSheet.Buttons.Add(Range("A2").Left, _
                                        Range("A2").Top, _
                                        Range("A2:B2").Width, _
                                        Range("A2:A4").Height)
                                            .name = "入力ボタン"
                                            .OnAction = "write_memo"
                                            .Characters.Text = "入力"
            End With
            'メモ削除ボタン作成
            With ActiveSheet.Buttons.Add(Range("A6").Left, _
                                        Range("A6").Top, _
                                        Range("A6:B6").Width, _
                                        Range("A6:A8").Height)
                                            .name = "削除ボタン"
                                            .OnAction = "close_memo"
                                            .Characters.Text = "閉じる"
            End With
            '時間割シート遷移ボタン作成
            With ActiveSheet.Buttons.Add(Range("A11").Left, _
                                        Range("A11").Top, _
                                        Range("A11:B11").Width, _
                                        Range("A11:A12").Height)
                                            .name = "ホームボタン"
                                            .OnAction = "go_home"
                                            .Characters.Text = "時間割"
            End With
            
            Range("A1").Value = target
            Range("A1").Font.Size = 30
            
            Range("C2").Value = "メモ欄："
            Range("C2").Font.Size = 22
            
            Range("E2").Value = "タグ・締切日"
            Range("E2").Font.Size = 18
            Range("E2").EntireColumn.AutoFit '横幅自動調整
            
            Range("F2").Value = "内容"
            Range("F2").Font.Size = 18
            
            Columns("F").ColumnWidth = 100 '横幅サイズ100
            
            Range("A9").Value = "メモ数："
            
            Range("B9").Value = 0
            
            Dim y As Long
            y = 3
            
            Do While Worksheets("Info").Cells(y, 7).Value <> ""
                y = y + 1
            Loop
            
            'Infoシートに登録
            Worksheets("Info").Cells(y, 7).Value = target
            
            ActiveWorkbook.Save '保存
            Exit Sub
        End If
        
        c_y = c_y + 1
    Loop
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox("アクティブセルを調整してください", Buttons:=vbCritical, Title:="エラー")
End Sub
Sub その他のメモ作成ボタン_Click()
    'メモのタイトルを入力させる
    Dim memo_name As String
    memo_name = InputBox("メモのタイトルを入力してください")
    
    '空欄だったらエラーを出して処理終了
    Dim ans As VbMsgBoxResult
    If memo_name = "" Then
        ans = MsgBox("名前を入力してください", Buttons:=vbExclamation, Title:="エラー")
        Exit Sub
    End If
    
    Dim c_y As Long
    c_y = 3
    
    '同じ名前のメモが存在するならエラーを出して終了
    Do While Worksheets("Info").Cells(c_y, 8).Value <> ""
        If memo_name = Worksheets("Info").Cells(c_y, 8).Value Then
            ans = MsgBox("タイトルが重複しています", Buttons:=vbCritical, Title:="エラー")
            Exit Sub
        End If
        
        c_y = c_y + 1
    Loop
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count)) 'シートを最後尾に追加
    
    'シート名
    ws.name = memo_name
    'タスク入力ボタン作成
    With ActiveSheet.Buttons.Add(Range("A2").Left, _
                                Range("A2").Top, _
                                Range("A2:B2").Width, _
                                Range("A2:A4").Height)
                                    .name = "入力ボタン"
                                    .OnAction = "write_memo"
                                    .Characters.Text = "入力"
    End With
    'メモ削除ボタン作成
    With ActiveSheet.Buttons.Add(Range("A6").Left, _
                                Range("A6").Top, _
                                Range("A6:B6").Width, _
                                Range("A6:A8").Height)
                                    .name = "削除ボタン"
                                    .OnAction = "close_memo"
                                    .Characters.Text = "閉じる"
    End With
    '時間割シート遷移ボタン作成
    With ActiveSheet.Buttons.Add(Range("A11").Left, _
                                Range("A11").Top, _
                                Range("A11:B11").Width, _
                                Range("A11:A12").Height)
                                    .name = "ホームボタン"
                                    .OnAction = "go_home"
                                    .Characters.Text = "時間割"
    End With
    
    Range("A1").Value = memo_name
    Range("A1").Font.Size = 30
    
    Range("C2").Value = "メモ欄："
    Range("C2").Font.Size = 22
    
    Range("E2").Value = "タグ・締切日"
    Range("E2").Font.Size = 18
    Range("E2").EntireColumn.AutoFit '横幅自動調整
    
    Range("F2").Value = "内容"
    Range("F2").Font.Size = 18
    
    Columns("F").ColumnWidth = 100 '横幅サイズ100
    
    Range("A9").Value = "メモ数："
    
    Range("B9").Value = 0
    
    Dim y As Long
    y = 3
    
    Do While Worksheets("Info").Cells(y, 8).Value <> ""
        y = y + 1
    Loop
    
    'Infoシートに登録
    Worksheets("Info").Cells(y, 8).Value = memo_name
    
    ActiveWorkbook.Save  '保存
End Sub
Sub 現在表示ボタン_Click()
    Dim day_today As Long '今日の曜日
    day_today = Weekday(Date) - 1
    Dim now_t As Date '現在時刻
    now_t = Time

    '背景色の初期化
    For x = 3 To 7
        For y = 3 To 8
            If y <> 5 Then
                Cells(y, x).Interior.ColorIndex = 0
            End If
        Next
    Next

    '現在の時刻に相当する授業の背景を黄色にする
    If 0 < day_today And day_today < 6 Then
        Select Case now_t
            Case CDate("8:50") To CDate("10:30")
                Cells(3, 2 + day_today).Interior.ColorIndex = 6
            Case CDate("10:45") To CDate("12:25")
                Cells(4, 2 + day_today).Interior.ColorIndex = 6
            Case CDate("13:15") To CDate("14:55")
                Cells(6, 2 + day_today).Interior.ColorIndex = 6
            Case CDate("15:10") To CDate("16:50")
                Cells(7, 2 + day_today).Interior.ColorIndex = 6
            Case CDate("17:05") To CDate("18:45")
                Cells(8, 2 + day_today).Interior.ColorIndex = 6
        End Select
    End If
End Sub
Sub フォルダ作成ボタン_Click()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim now_path As String
    now_path = GetPath(ActiveWorkbook.Path) '現在のパス
    Dim target
    Set target = fso.GetFolder(now_path).SubFolders
    Dim tmp As Object
    
    '重複するフォルダが作られていなければ各教科のフォルダを新規作成
    For x = 3 To 7
        For y = 3 To 8
            If Cells(y, x) <> "" Then
                If Dir(now_path + "\" + Cells(y, x).Value, vbDirectory) = "" Then
                    MkDir (now_path + "\" + Cells(y, x).Value)
                End If
            End If
        Next
    Next
End Sub
Sub 更新ボタン_Click()
    '時間割シートのフォント設定を初期化
    For x = 3 To 7
        For y = 3 To 8
            If Worksheets("時間割").Cells(y, x).Value <> "" Then
                Worksheets("時間割").Cells(y, x).Font.Color = 1
                Worksheets("時間割").Cells(y, x).Font.Bold = False
            End If
        Next
    Next
    
    Dim todo_y As Long
    todo_y = 3
    
    'Infoシートの初期化
    Do While Worksheets("Info").Cells(todo_y, 9).Value <> "" Or Worksheets("Info").Cells(todo_y, 10).Value <> ""
        Worksheets("Info").Cells(todo_y, 9).Value = ""
        Worksheets("Info").Cells(todo_y, 10).Value = ""
        
        todo_y = todo_y + 1
    Loop
    
    Dim today As Date
    today = Date '今日
    Dim tomorrow As Date
    tomorrow = DateAdd("d", 1, Date) '明日
    
    For Each ws In Worksheets
        Dim c_y As Long
        c_y = 3
        
        Dim sbj_x As Long
        Dim sbj_y As Long
        
        '時間割シートとInfoシート以外に
        If ws.name <> "時間割" Or ws.name <> "Info" Then
            Do While Worksheets(ws.name).Cells(c_y, 6).Value <> ""
                'シート内の締切日が
                Select Case Worksheets(ws.name).Cells(c_y, 5).Value
                    '今日なら
                    Case today
                        '対応するものを赤に色付けする
                        Worksheets(ws.name).Cells(c_y, 5).Interior.ColorIndex = 3
                        'シートを時間割の左隣に
                        Call Worksheets(ws.name).Move(After:=Sheets("時間割"))
                        
                        todo_y = 3
                        
                        Do While Worksheets("Info").Cells(todo_y, 2).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 2).Value = ws.name Then
                                sbj_x = Worksheets("Info").Cells(todo_y, 4).Value
                                sbj_y = Worksheets("Info").Cells(todo_y, 5).Value
                                
                                '時間割表の対応する教科を
                                '赤に色付けする
                                Worksheets("時間割").Cells(sbj_y, sbj_x).Font.ColorIndex = 3
                                '太字にする
                                Worksheets("時間割").Cells(sbj_y, sbj_x).Font.Bold = True
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        
                        todo_y = 3
                        
                        'Infoシートに登録
                        Do While Worksheets("Info").Cells(todo_y, 9).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 9).Value = ws.name Then
                                Exit Do
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        Worksheets("Info").Cells(todo_y, 9).Value = ws.name
                    '明日なら
                    Case tomorrow
                        '対応するものをピンクに色付けする
                        Worksheets(ws.name).Cells(c_y, 5).Interior.ColorIndex = 38
                        
                        todo_y = 3
                        
                        Do While Worksheets("Info").Cells(todo_y, 2).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 2).Value = ws.name Then
                                sbj_x = Worksheets("Info").Cells(todo_y, 4).Value
                                sbj_y = Worksheets("Info").Cells(todo_y, 5).Value
                                
                                '時間割表の対応する教科が赤（明日まで）に色付けされていないなら
                                If Worksheets("時間割").Cells(sbj_y, sbj_x).Font.ColorIndex <> 3 Then
                                    '時間割表に対応する教科を
                                    'ピンクに色付けする
                                    Worksheets("時間割").Cells(sbj_y, sbj_x).Font.ColorIndex = 38
                                    '太字にする
                                    Worksheets("時間割").Cells(sbj_y, sbj_x).Font.Bold = True
                                End If
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        
                        todo_y = 3
                        
                        'Infoシートに登録
                        Do While Worksheets("Info").Cells(todo_y, 10).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 10).Value = ws.name Then
                                Exit Do
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        Worksheets("Info").Cells(todo_y, 10).Value = ws.name
                    'それ以外なら
                    Case Else
                        '色付けを初期化
                        Worksheets(ws.name).Cells(c_y, 5).Interior.ColorIndex = 0
                End Select
                
                c_y = c_y + 1
            Loop
        End If
    Next ws
    
    Worksheets("時間割").Activate
    ActiveWorkbook.Save
End Sub
Sub メモ一覧_Click() 'メモ一覧の表示
    'メモ一覧シートを作成
    Dim ws As Worksheet
    Set ws = Sheets.Add
    ws.name = "メモ一覧"
    ws.OnSheetDeactivate = "ws_delete" 'シートを離れたら削除
    
    '移動ボタン作成
    With ActiveSheet.Buttons.Add(Range("A1").Left, _
                                Range("A1").Top, _
                                Range("A1:B1").Width, _
                                Range("A1:A2").Height)
                                    .name = "移動ボタン"
                                    .OnAction = "change"
                                    .Characters.Text = "移動"
    End With
    
    Dim c_y As Long
    c_y = 4
    
    'それぞれのワークシート名で判定
    For Each ws In Worksheets
        If ws.name <> "時間割" And ws.name <> "Info" And ws.name <> "メモ一覧" Then
            Dim y As Long
            y = 3
            
            Do While True
                '教科のメモなら
                If Worksheets("Info").Cells(y, 7).Value = ws.name Then
                    'メモ一覧シートに載せる
                    Worksheets("メモ一覧").Cells(c_y, 1).Value = ws.name
                    '緑色の背景にする
                    Worksheets("メモ一覧").Cells(c_y, 1).Interior.ColorIndex = 4
                    
                    Exit Do
                'その他のメモなら
                ElseIf Worksheets("Info").Cells(y, 8).Value = ws.name Then
                    'メモ一覧シートに載せる
                    Worksheets("メモ一覧").Cells(c_y, 1).Value = ws.name
                    '水色の背景にする
                    Worksheets("メモ一覧").Cells(c_y, 1).Interior.ColorIndex = 8
                    
                    Exit Do
                Else
                    y = y + 1
                End If
            Loop
            
            c_y = c_y + 1
        End If
    Next ws
    
    'セル全体を自動調整する
    Cells.EntireColumn.AutoFit
End Sub
Sub 通知ボタン_Click() 'マイパレットの呼び出し
    Shell Worksheets("Info").Cells(3, 12), vbNormalFocus
    ActiveWorkbook.Save
End Sub
Sub Info表示_Click()
    Dim ans As String
    ans = InputBox("パスワードを入力してください")
    
    'Infoシートに設定したパスワードが合っていれば表示
    If ans = Worksheets("Info").Cells(2, 11).Value Then
        Worksheets("Info").Visible = xlSheetVisible
        Worksheets("Info").Activate
    Else
        ans = MsgBox("パスワードが違います", Buttons:=vbCritical, Title:="エラー")
    End If
End Sub

