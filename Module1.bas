Attribute VB_Name = "Module1"
Sub ���Ȋm��{�^��_Click()
    '���Ȗ��o�^�Ȃ�
    Dim ans As VbMsgBoxResult
    
    If Worksheets("Info").Cells(2, 2).Value = 0 Then
        ans = MsgBox("���Ȃ��m�肵�܂����H", Buttons:=vbYesNo + vbQuestion, Title:="�m�F")
        
        If ans = vbYes Then
            Dim c_y As Long
            c_y = 3
            
            For x = 3 To 7
                For y = 3 To 8
                    '���Ԋ��\���狳�ȂƑΉ�������W��Info�V�[�g�ɓo�^
                    If Worksheets("���Ԋ�").Cells(y, x).Value <> "" Then
                        Worksheets("Info").Cells(c_y, 2).Value = Worksheets("���Ԋ�").Cells(y, x).Value
                        Worksheets("Info").Cells(c_y, 4).Value = x
                        Worksheets("Info").Cells(c_y, 5).Value = y
                        
                        c_y = c_y + 1
                    End If
                Next
            Next
            
            ans = MsgBox("���Ȃ̓o�^���������܂���" & vbCrLf & "�����̓o�^�����{���Ă�������", Buttons:=vbInformation, Title:="�o�^")
            '���Ȃ�o�^��Ԃɂ���
            Worksheets("Info").Cells(2, 2).Value = 1
            
            '�ۑ�
            ActiveWorkbook.Save
        End If
    Else
        ans = MsgBox("���Ȃ͊��ɓo�^����Ă��܂�", Buttons:=vbCritical, Title:="�o�^��")
    End If
End Sub
Sub �����m��{�^��_Click()
    Dim ans As VbMsgBoxResult
    
    '���Ȃ��o�^��Ԃŋ��������o�^�Ȃ�
    If Worksheets("Info").Cells(2, 3).Value = 0 And Worksheets("Info").Cells(2, 2).Value = 1 Then
        Dim c_y As Long
        c_y = 3
        
        Do While Worksheets("Info").Cells(c_y, 2).Value <> ""
            '�Ή����鋳�Ȃ̎��Ԋ��V�[�g�ɂ����Ă�xy���W
            Dim x As Long
            x = Worksheets("Info").Cells(c_y, 4).Value
            Dim y As Long
            y = Worksheets("Info").Cells(c_y, 5).Value
            '�Ή�����xy���W�̔w�i�F�����F�ɂ���
            Worksheets("���Ԋ�").Cells(y, x).Interior.ColorIndex = 6
            
            classroom = InputBox(Worksheets("Info").Cells(c_y, 2).Value & "�̋�������͂��Ă�������")
            Worksheets("Info").Cells(c_y, 3).Value = classroom
            
            '�w�i�F�����Ƃɖ߂�
            Worksheets("���Ԋ�").Cells(y, x).Interior.ColorIndex = 0
                        
            c_y = c_y + 1
        Loop
        
        ans = MsgBox("�����̓o�^���������܂���", Buttons:=vbInformation, Title:="�o�^")
        '������o�^��Ԃɂ���
        Worksheets("Info").Cells(2, 3).Value = 1
        
        '�ۑ�
        ActiveWorkbook.Save
    '�������o�^�Ȃ�
    ElseIf Worksheets("Info").Cells(2, 2).Value = 0 Then
        ans = MsgBox("��ɋ��Ȃ̓o�^�����Ă�������", Buttons:=vbCritical, Title:="�G���[")
    Else
        ans = MsgBox("�����͊��ɓo�^����Ă��܂�", Buttons:=vbCritical, Title:="�o�^��")
    End If
End Sub
Sub ������_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("���������܂����H", Buttons:=vbYesNo + vbExclamation, Title:="�m�F")
    
    If ans = vbYes Then
        '�o�^��Ԃ̏�����
        Worksheets("Info").Cells(2, 2).Value = 0
        Worksheets("Info").Cells(2, 3).Value = 0
        
        Dim c_y As Long
        c_y = 3
        
        For x = 3 To 7
            For y = 3 To 8
                '�w�i�F�̏�����
                Worksheets("���Ԋ�").Cells(y, x).Interior.ColorIndex = 0
                
                If Worksheets("���Ԋ�").Cells(y, x).Value <> "" Then
                    '���Ԋ��V�[�g�̋��Ȃ�������
                    Worksheets("���Ԋ�").Cells(y, x).Value = ""
                    
                    'Info�V�[�g�̏���������
                    Worksheets("Info").Cells(c_y, 2).Value = ""
                    Worksheets("Info").Cells(c_y, 3).Value = ""
                    Worksheets("Info").Cells(c_y, 4).Value = ""
                    Worksheets("Info").Cells(c_y, 5).Value = ""
                        
                    c_y = c_y + 1
                End If
            Next
        Next
        
        '�ۑ�
        ActiveWorkbook.Save
    End If
End Sub
Sub ���Ȑ؂�ւ��{�^��_Click()
    '�o�^��ԂȂ�
    If Worksheets("Info").Cells(2, 2).Value = 1 And Worksheets("Info").Cells(2, 3).Value = 1 Then
        Dim c_y As Long
        Dim x As Long
        Dim y As Long
        
        '���݂����ȕ\����ԂȂ�
        If Worksheets("Info").Cells(2, 6).Value = "sbj" Then
            c_y = 3
            
            Do While Worksheets("Info").Cells(c_y, 2).Value <> ""
                x = Worksheets("Info").Cells(c_y, 4).Value
                y = Worksheets("Info").Cells(c_y, 5).Value
                
                '���Ԋ��V�[�g�̑Ή����鋳�Ȃ������ɂ���
                Worksheets("���Ԋ�").Cells(y, x).Value = Worksheets("Info").Cells(c_y, 3).Value
                c_y = c_y + 1
            Loop
            '�����\����Ԃɂ���
            Worksheets("Info").Cells(2, 6).Value = "cls"
        '���݂������\����ԂȂ�
        ElseIf Worksheets("Info").Cells(2, 6).Value = "cls" Then
            c_y = 3
            
            Do While Worksheets("Info").Cells(c_y, 3).Value <> ""
                x = Worksheets("Info").Cells(c_y, 4).Value
                y = Worksheets("Info").Cells(c_y, 5).Value
                
                '���Ԋ��V�[�g�̑Ή����鋳�������Ȃɂ���
                Worksheets("���Ԋ�").Cells(y, x).Value = Worksheets("Info").Cells(c_y, 2).Value
                c_y = c_y + 1
            Loop
            '���ȕ\����Ԃɂ���
            Worksheets("Info").Cells(2, 6).Value = "sbj"
        End If
    Else
        Dim ans As VbMsgBoxResult
        ans = MsgBox("���ȁE�����o�^���������Ă�������", Buttons:=vbCritical, Title:="�G���[")
    End If
End Sub
Sub �����쐬�{�^��_Click() '���ȗp�����쐬
    '�A�N�e�B�u�Z����ǂݍ���
    Dim target As String
    target = ActiveCell.Value
    
    Dim c_y As Long
    c_y = 3
    
    '�������w�肳��Ă���Ȃ狳�Ȃɕϊ�����
    Do While Worksheets("Info").Cells(c_y, 3).Value <> ""
        If target = Worksheets("Info").Cells(c_y, 3).Value Then
            target = Worksheets("Info").Cells(c_y, 2).Value
            Exit Do
        End If
        
        c_y = c_y + 1
    Loop
    
    '���������ɍ쐬����Ă���Ȃ炻����A�N�e�B�u�ɂ��A�������I������
    For Each memo In Worksheets
        If memo.name = target Then
            memo.Activate
            Exit Sub
        End If
    Next memo
    
    c_y = 3
    
    'Info�V�[�g�ɓo�^���ꂽ���Ȃƈ�v����܂�
    Do While Worksheets("Info").Cells(c_y, 2).Value <> ""
        '�A�N�e�B�u�Z�������ȂƋ����̂ǂ��炩����v����Ȃ�
        If Worksheets("Info").Cells(c_y, 2).Value = target Or Worksheets("Info").Cells(c_y, 3).Value = target Then
            Dim ws As Worksheet
            Set ws = Sheets.Add(After:=Sheets(Sheets.Count)) '�V�[�g���Ō���ɒǉ�
            
            '�V�[�g��
            ws.name = target
            '�^�X�N���̓{�^���쐬
            With ActiveSheet.Buttons.Add(Range("A2").Left, _
                                        Range("A2").Top, _
                                        Range("A2:B2").Width, _
                                        Range("A2:A4").Height)
                                            .name = "���̓{�^��"
                                            .OnAction = "write_memo"
                                            .Characters.Text = "����"
            End With
            '�����폜�{�^���쐬
            With ActiveSheet.Buttons.Add(Range("A6").Left, _
                                        Range("A6").Top, _
                                        Range("A6:B6").Width, _
                                        Range("A6:A8").Height)
                                            .name = "�폜�{�^��"
                                            .OnAction = "close_memo"
                                            .Characters.Text = "����"
            End With
            '���Ԋ��V�[�g�J�ڃ{�^���쐬
            With ActiveSheet.Buttons.Add(Range("A11").Left, _
                                        Range("A11").Top, _
                                        Range("A11:B11").Width, _
                                        Range("A11:A12").Height)
                                            .name = "�z�[���{�^��"
                                            .OnAction = "go_home"
                                            .Characters.Text = "���Ԋ�"
            End With
            
            Range("A1").Value = target
            Range("A1").Font.Size = 30
            
            Range("C2").Value = "�������F"
            Range("C2").Font.Size = 22
            
            Range("E2").Value = "�^�O�E���ؓ�"
            Range("E2").Font.Size = 18
            Range("E2").EntireColumn.AutoFit '������������
            
            Range("F2").Value = "���e"
            Range("F2").Font.Size = 18
            
            Columns("F").ColumnWidth = 100 '�����T�C�Y100
            
            Range("A9").Value = "�������F"
            
            Range("B9").Value = 0
            
            Dim y As Long
            y = 3
            
            Do While Worksheets("Info").Cells(y, 7).Value <> ""
                y = y + 1
            Loop
            
            'Info�V�[�g�ɓo�^
            Worksheets("Info").Cells(y, 7).Value = target
            
            ActiveWorkbook.Save '�ۑ�
            Exit Sub
        End If
        
        c_y = c_y + 1
    Loop
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox("�A�N�e�B�u�Z���𒲐����Ă�������", Buttons:=vbCritical, Title:="�G���[")
End Sub
Sub ���̑��̃����쐬�{�^��_Click()
    '�����̃^�C�g������͂�����
    Dim memo_name As String
    memo_name = InputBox("�����̃^�C�g������͂��Ă�������")
    
    '�󗓂�������G���[���o���ď����I��
    Dim ans As VbMsgBoxResult
    If memo_name = "" Then
        ans = MsgBox("���O����͂��Ă�������", Buttons:=vbExclamation, Title:="�G���[")
        Exit Sub
    End If
    
    Dim c_y As Long
    c_y = 3
    
    '�������O�̃��������݂���Ȃ�G���[���o���ďI��
    Do While Worksheets("Info").Cells(c_y, 8).Value <> ""
        If memo_name = Worksheets("Info").Cells(c_y, 8).Value Then
            ans = MsgBox("�^�C�g�����d�����Ă��܂�", Buttons:=vbCritical, Title:="�G���[")
            Exit Sub
        End If
        
        c_y = c_y + 1
    Loop
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count)) '�V�[�g���Ō���ɒǉ�
    
    '�V�[�g��
    ws.name = memo_name
    '�^�X�N���̓{�^���쐬
    With ActiveSheet.Buttons.Add(Range("A2").Left, _
                                Range("A2").Top, _
                                Range("A2:B2").Width, _
                                Range("A2:A4").Height)
                                    .name = "���̓{�^��"
                                    .OnAction = "write_memo"
                                    .Characters.Text = "����"
    End With
    '�����폜�{�^���쐬
    With ActiveSheet.Buttons.Add(Range("A6").Left, _
                                Range("A6").Top, _
                                Range("A6:B6").Width, _
                                Range("A6:A8").Height)
                                    .name = "�폜�{�^��"
                                    .OnAction = "close_memo"
                                    .Characters.Text = "����"
    End With
    '���Ԋ��V�[�g�J�ڃ{�^���쐬
    With ActiveSheet.Buttons.Add(Range("A11").Left, _
                                Range("A11").Top, _
                                Range("A11:B11").Width, _
                                Range("A11:A12").Height)
                                    .name = "�z�[���{�^��"
                                    .OnAction = "go_home"
                                    .Characters.Text = "���Ԋ�"
    End With
    
    Range("A1").Value = memo_name
    Range("A1").Font.Size = 30
    
    Range("C2").Value = "�������F"
    Range("C2").Font.Size = 22
    
    Range("E2").Value = "�^�O�E���ؓ�"
    Range("E2").Font.Size = 18
    Range("E2").EntireColumn.AutoFit '������������
    
    Range("F2").Value = "���e"
    Range("F2").Font.Size = 18
    
    Columns("F").ColumnWidth = 100 '�����T�C�Y100
    
    Range("A9").Value = "�������F"
    
    Range("B9").Value = 0
    
    Dim y As Long
    y = 3
    
    Do While Worksheets("Info").Cells(y, 8).Value <> ""
        y = y + 1
    Loop
    
    'Info�V�[�g�ɓo�^
    Worksheets("Info").Cells(y, 8).Value = memo_name
    
    ActiveWorkbook.Save  '�ۑ�
End Sub
Sub ���ݕ\���{�^��_Click()
    Dim day_today As Long '�����̗j��
    day_today = Weekday(Date) - 1
    Dim now_t As Date '���ݎ���
    now_t = Time

    '�w�i�F�̏�����
    For x = 3 To 7
        For y = 3 To 8
            If y <> 5 Then
                Cells(y, x).Interior.ColorIndex = 0
            End If
        Next
    Next

    '���݂̎����ɑ���������Ƃ̔w�i�����F�ɂ���
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
Sub �t�H���_�쐬�{�^��_Click()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim now_path As String
    now_path = GetPath(ActiveWorkbook.Path) '���݂̃p�X
    Dim target
    Set target = fso.GetFolder(now_path).SubFolders
    Dim tmp As Object
    
    '�d������t�H���_������Ă��Ȃ���Ίe���Ȃ̃t�H���_��V�K�쐬
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
Sub �X�V�{�^��_Click()
    '���Ԋ��V�[�g�̃t�H���g�ݒ��������
    For x = 3 To 7
        For y = 3 To 8
            If Worksheets("���Ԋ�").Cells(y, x).Value <> "" Then
                Worksheets("���Ԋ�").Cells(y, x).Font.Color = 1
                Worksheets("���Ԋ�").Cells(y, x).Font.Bold = False
            End If
        Next
    Next
    
    Dim todo_y As Long
    todo_y = 3
    
    'Info�V�[�g�̏�����
    Do While Worksheets("Info").Cells(todo_y, 9).Value <> "" Or Worksheets("Info").Cells(todo_y, 10).Value <> ""
        Worksheets("Info").Cells(todo_y, 9).Value = ""
        Worksheets("Info").Cells(todo_y, 10).Value = ""
        
        todo_y = todo_y + 1
    Loop
    
    Dim today As Date
    today = Date '����
    Dim tomorrow As Date
    tomorrow = DateAdd("d", 1, Date) '����
    
    For Each ws In Worksheets
        Dim c_y As Long
        c_y = 3
        
        Dim sbj_x As Long
        Dim sbj_y As Long
        
        '���Ԋ��V�[�g��Info�V�[�g�ȊO��
        If ws.name <> "���Ԋ�" Or ws.name <> "Info" Then
            Do While Worksheets(ws.name).Cells(c_y, 6).Value <> ""
                '�V�[�g���̒��ؓ���
                Select Case Worksheets(ws.name).Cells(c_y, 5).Value
                    '�����Ȃ�
                    Case today
                        '�Ή�������̂�ԂɐF�t������
                        Worksheets(ws.name).Cells(c_y, 5).Interior.ColorIndex = 3
                        '�V�[�g�����Ԋ��̍��ׂ�
                        Call Worksheets(ws.name).Move(After:=Sheets("���Ԋ�"))
                        
                        todo_y = 3
                        
                        Do While Worksheets("Info").Cells(todo_y, 2).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 2).Value = ws.name Then
                                sbj_x = Worksheets("Info").Cells(todo_y, 4).Value
                                sbj_y = Worksheets("Info").Cells(todo_y, 5).Value
                                
                                '���Ԋ��\�̑Ή����鋳�Ȃ�
                                '�ԂɐF�t������
                                Worksheets("���Ԋ�").Cells(sbj_y, sbj_x).Font.ColorIndex = 3
                                '�����ɂ���
                                Worksheets("���Ԋ�").Cells(sbj_y, sbj_x).Font.Bold = True
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        
                        todo_y = 3
                        
                        'Info�V�[�g�ɓo�^
                        Do While Worksheets("Info").Cells(todo_y, 9).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 9).Value = ws.name Then
                                Exit Do
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        Worksheets("Info").Cells(todo_y, 9).Value = ws.name
                    '�����Ȃ�
                    Case tomorrow
                        '�Ή�������̂��s���N�ɐF�t������
                        Worksheets(ws.name).Cells(c_y, 5).Interior.ColorIndex = 38
                        
                        todo_y = 3
                        
                        Do While Worksheets("Info").Cells(todo_y, 2).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 2).Value = ws.name Then
                                sbj_x = Worksheets("Info").Cells(todo_y, 4).Value
                                sbj_y = Worksheets("Info").Cells(todo_y, 5).Value
                                
                                '���Ԋ��\�̑Ή����鋳�Ȃ��ԁi�����܂Łj�ɐF�t������Ă��Ȃ��Ȃ�
                                If Worksheets("���Ԋ�").Cells(sbj_y, sbj_x).Font.ColorIndex <> 3 Then
                                    '���Ԋ��\�ɑΉ����鋳�Ȃ�
                                    '�s���N�ɐF�t������
                                    Worksheets("���Ԋ�").Cells(sbj_y, sbj_x).Font.ColorIndex = 38
                                    '�����ɂ���
                                    Worksheets("���Ԋ�").Cells(sbj_y, sbj_x).Font.Bold = True
                                End If
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        
                        todo_y = 3
                        
                        'Info�V�[�g�ɓo�^
                        Do While Worksheets("Info").Cells(todo_y, 10).Value <> ""
                            If Worksheets("Info").Cells(todo_y, 10).Value = ws.name Then
                                Exit Do
                            End If
                            
                            todo_y = todo_y + 1
                        Loop
                        Worksheets("Info").Cells(todo_y, 10).Value = ws.name
                    '����ȊO�Ȃ�
                    Case Else
                        '�F�t����������
                        Worksheets(ws.name).Cells(c_y, 5).Interior.ColorIndex = 0
                End Select
                
                c_y = c_y + 1
            Loop
        End If
    Next ws
    
    Worksheets("���Ԋ�").Activate
    ActiveWorkbook.Save
End Sub
Sub �����ꗗ_Click() '�����ꗗ�̕\��
    '�����ꗗ�V�[�g���쐬
    Dim ws As Worksheet
    Set ws = Sheets.Add
    ws.name = "�����ꗗ"
    ws.OnSheetDeactivate = "ws_delete" '�V�[�g�𗣂ꂽ��폜
    
    '�ړ��{�^���쐬
    With ActiveSheet.Buttons.Add(Range("A1").Left, _
                                Range("A1").Top, _
                                Range("A1:B1").Width, _
                                Range("A1:A2").Height)
                                    .name = "�ړ��{�^��"
                                    .OnAction = "change"
                                    .Characters.Text = "�ړ�"
    End With
    
    Dim c_y As Long
    c_y = 4
    
    '���ꂼ��̃��[�N�V�[�g���Ŕ���
    For Each ws In Worksheets
        If ws.name <> "���Ԋ�" And ws.name <> "Info" And ws.name <> "�����ꗗ" Then
            Dim y As Long
            y = 3
            
            Do While True
                '���Ȃ̃����Ȃ�
                If Worksheets("Info").Cells(y, 7).Value = ws.name Then
                    '�����ꗗ�V�[�g�ɍڂ���
                    Worksheets("�����ꗗ").Cells(c_y, 1).Value = ws.name
                    '�ΐF�̔w�i�ɂ���
                    Worksheets("�����ꗗ").Cells(c_y, 1).Interior.ColorIndex = 4
                    
                    Exit Do
                '���̑��̃����Ȃ�
                ElseIf Worksheets("Info").Cells(y, 8).Value = ws.name Then
                    '�����ꗗ�V�[�g�ɍڂ���
                    Worksheets("�����ꗗ").Cells(c_y, 1).Value = ws.name
                    '���F�̔w�i�ɂ���
                    Worksheets("�����ꗗ").Cells(c_y, 1).Interior.ColorIndex = 8
                    
                    Exit Do
                Else
                    y = y + 1
                End If
            Loop
            
            c_y = c_y + 1
        End If
    Next ws
    
    '�Z���S�̂�������������
    Cells.EntireColumn.AutoFit
End Sub
Sub �ʒm�{�^��_Click() '�}�C�p���b�g�̌Ăяo��
    Shell Worksheets("Info").Cells(3, 12), vbNormalFocus
    ActiveWorkbook.Save
End Sub
Sub Info�\��_Click()
    Dim ans As String
    ans = InputBox("�p�X���[�h����͂��Ă�������")
    
    'Info�V�[�g�ɐݒ肵���p�X���[�h�������Ă���Ε\��
    If ans = Worksheets("Info").Cells(2, 11).Value Then
        Worksheets("Info").Visible = xlSheetVisible
        Worksheets("Info").Activate
    Else
        ans = MsgBox("�p�X���[�h���Ⴂ�܂�", Buttons:=vbCritical, Title:="�G���[")
    End If
End Sub

