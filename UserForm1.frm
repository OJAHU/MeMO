VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����"
   ClientHeight    =   4440
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7080
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
    
    '���e��������Ă���Ȃ�
    If TextBox2.Text <> "" Then
        '���ؓ������
        If TextBox1.Text = "����" Then
            Cells(c_y, 5).Value = Format(Date, "mm/dd")
            Cells(c_y, 6).Value = TextBox2.Text
        ElseIf TextBox1.Value = "����" Then
            Cells(c_y, 5).Value = DateAdd("d", 1, Date)
        ElseIf TextBox1.Value = "���T" Then
            Cells(c_y, 5).Value = DateAdd("d", 7, Date)
        Else
            Cells(c_y, 5).Value = TextBox1.Text
        End If
        
        '���e�����
        Cells(c_y, 6).Value = TextBox2.Text
        
        '�e�^�X�N�̏I���{�^���쐬
        With ActiveSheet.Buttons.Add(Cells(c_y, 4).Left, _
                                    Cells(c_y, 4).Top, _
                                    Cells(c_y, 4).Width, _
                                    Cells(c_y, 4).Height)
                                        .name = c_y
                                        .OnAction = "fin_task"
                                        .Characters.Text = "�I��"
        End With
        
        '�������𑝂₷
        Range("B9").Value = Range("B9").Value + 1
        
        ActiveWorkbook.Save
        Unload Me
    Else
        Dim ans As String
        ans = MsgBox("���e����͂��Ă�������", Buttons:=vbOKOnly, Title:="�G���[")
    End If
End Sub

