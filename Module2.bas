Attribute VB_Name = "Module2"
Sub write_memo() '�^�X�N�쐬
    Load UserForm1 '���[�U�[�t�H�[���̌Ăяo��
    UserForm1.Show
End Sub
Sub close_memo() '���������
    '��������0�Ȃ�
    If Range("B9").Value = 0 Then
        Dim name As String
        name = Range("A1").Value
            
        '�폜�E�C���h�E
        Dim ans As Variant
        ans = ActiveSheet.Delete
        
        '�E�C���h�E�ō폜����������
        If ans = True Then
            Dim c_y As Long
            c_y = 3
            
            '������
        
            Dim y As Long
            
            Do While Worksheets("Info").Cells(c_y, 7).Value <> "" Or Worksheets("Info").Cells(c_y, 8).Value <> ""
                '���Ȃ̃����Ȃ�
                If name = Worksheets("Info").Cells(c_y, 7).Value Then
                    'Info�V�[�g�̑Ή�������̂��󔒂ɂ���
                    Worksheets("Info").Cells(c_y, 7).Value = ""
                    
                    y = c_y + 1
                    
                    'Info�V�[�g�̋󗓕������Ȃ����悤�ɂ��炷
                    Do While Worksheets("Info").Cells(y, 7).Value <> ""
                        Worksheets("Info").Cells(y - 1, 7).Value = Worksheets("Info").Cells(y, 7).Value
                        Worksheets("Info").Cells(y, 7).Value = ""
                        
                        y = y + 1
                    Loop
                    
                    Exit Do
                    
                    
                '���̑��̃����Ȃ�
                ElseIf name = Worksheets("Info").Cells(c_y, 8).Value Then
                    'Info�V�[�g�̑Ή�������̂��󔒂ɂ���
                    Worksheets("Info").Cells(c_y, 8).Value = ""
                    
                    y = c_y + 1
                    
                    'Info�V�[�g�̋󗓕������Ȃ����悤�ɂ��炷
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
        rslt = MsgBox("�^�X�N�����ׂďI�����Ă�����Ă�������", Buttons:=vbCritical, Title:="�G���[")
    End If
    
    ActiveWorkbook.Save
End Sub
Sub go_home() '���Ԋ��V�[�g�ɖ߂�
    Worksheets("���Ԋ�").Activate
End Sub
Sub fin_task() '������̃^�X�N���폜
    Dim rslt As VbMsgBoxResult
    rslt = MsgBox("�^�X�N���I�����܂����H", Buttons:=vbYesNo, Title:="�I��")
    
    If rslt = vbYes Then
        Dim btn As Object
        Dim pushed As Long '�����ꂽ�{�^����y���W���擾
        pushed = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row
        
        Range(Cells(pushed, 5), Cells(pushed, 6)).Value = ""
        Range(Cells(pushed, 5), Cells(pushed, 6)).Interior.ColorIndex = 0
           
        '�I���{�^�����폜
        For Each btn In ActiveSheet.Buttons
            If IsNumeric(btn.name) Then
                btn.Delete
            End If
        Next btn
        
        Dim c_y As Long
        c_y = pushed + 1
        
        '�󔒕��������炵���߂�
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
                                                .Characters.Text = "�I��"
            End With
            c_y = c_y + 1
        Loop
        
        Range("B9").Value = Range("B9").Value - 1
    End If
End Sub
Sub change() '�����̈ړ�
    Dim subject As String
    subject = ActiveCell
    
    '���Ƀ������쐬����Ă���Ȃ炻����J��
    On Error GoTo ErrLabel
    Worksheets(subject).Activate
    
    '�����ꗗ�̍폜
    Application.DisplayAlerts = False
    Worksheets("�����ꗗ").Delete
    Application.DisplayAlerts = True
    Exit Sub
    
ErrLabel:
    Dim ans As VbMsgBoxResult
    ans = MsgBox("�A�N�e�B�u�Z���𒲐����Ă�������", Buttons:=vbCritical, Title:="�G���[")
End Sub
Sub ws_delete()
    '�����ꗗ�̍폜
    Application.DisplayAlerts = False
    Worksheets("�����ꗗ").Delete
    Application.DisplayAlerts = True
End Sub


