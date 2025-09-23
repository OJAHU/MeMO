Attribute VB_Name = "Module3"
' OneDrive���̃p�X�̎擾����url�ɂȂ�\��������p�X�����[�J���p�X�ɕϊ�
Function GetPath(url As String) As String
    Dim oneDrivePath As String
    Dim userName As String
    Dim shortcut As Object
    Dim subfolder As Object
    Dim fso As Object  ' �t�@�C���V�X�e����������邽�߂̃I�u�W�F�N�g
    
    userName = Environ("USERNAME")
    oneDrivePath = "C:\Users\" & userName & Worksheets("Info").Cells(3, 15).Value
    
    Dim folder As Object
    ' FileSystemObject�̃C���X�^���X���쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(oneDrivePath)
    
    '�p�X��https:���܂܂��ꍇ
    If url Like "https:*" Then
        For Each subfolder In folder.SubFolders
            Dim subfolderName As String
            subfolderName = subfolder.name

            Dim subfolderPosition As Integer
            subfolderPosition = InStr(1, url, subfolderName & "/") ' subfolderName�̈ʒu���擾

            If subfolderPosition > 0 Then
                Dim relativePath As String
                relativePath = Mid(url, subfolderPosition) ' subfolderName���܂ޕ�������̃p�X���擾
                GetPath = oneDrivePath & relativePath ' OneDrive�p�X�ɂ��̕�����ǉ�
                Exit Function
            End If
        Next subfolder
    Else
        GetPath = url 'https���܂܂Ȃ��ꍇ�͂��̂܂ܕԂ�
    End If
End Function
