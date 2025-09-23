Attribute VB_Name = "Module4"
Sub make_folder()
    'Myfolder作成
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'もし同じ物がないなら作成
    If Dir(Worksheets("Info").Cells(3, 13).Value, vbDirectory) = "" Then
        MkDir (Worksheets("Info").Cells(3, 13).Value)
    End If
End Sub
Sub teller()
    Call make_folder
    Call 更新ボタン_Click
    
    'ファイルがなければ作成
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.fileexists(Worksheets("Info").Cells(3, 14).Value) Then
        fso.createtextfile (Worksheets("Info").Cells(3, 14).Value)
    End If
    
    Dim c_y As Long
    Dim task As Long
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        
        c_y = 3
        
        Do While Worksheets("Info").Cells(c_y, 9) <> ""
            ws = Worksheets("Info").Cells(c_y, 9).Value
            
            task = 3
            
            Do While Worksheets(ws).Cells(task, 6) <> ""
                If Worksheets(ws).Cells(task, 5).Value = Date Then
                    'Block Of Task
                    .writetext "0," + ws + "," + Worksheets(ws).Cells(task, 6).Value + "BOT"
                End If
                
                task = task + 1
            Loop
            
            c_y = c_y + 1
        Loop

        c_y = 3
        
        Do While Worksheets("Info").Cells(c_y, 10) <> ""
            ws = Worksheets("Info").Cells(c_y, 10).Value
            
            task = 3
            
            Do While Worksheets(ws).Cells(task, 6) <> ""
                If Worksheets(ws).Cells(task, 5).Value = Date + 1 Then
                    'Block Of Task
                    .writetext "1," + ws + "," + Worksheets(ws).Cells(task, 6).Value + "BOT"
                End If
                
                task = task + 1
            Loop
            
            c_y = c_y + 1
        Loop

        .savetofile Worksheets("Info").Cells(3, 14).Value, 2
        .Close
    End With
End Sub
