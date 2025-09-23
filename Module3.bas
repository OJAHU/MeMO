Attribute VB_Name = "Module3"
' OneDrive内のパスの取得時にurlになる可能性があるパスをローカルパスに変換
Function GetPath(url As String) As String
    Dim oneDrivePath As String
    Dim userName As String
    Dim shortcut As Object
    Dim subfolder As Object
    Dim fso As Object  ' ファイルシステム操作をするためのオブジェクト
    
    userName = Environ("USERNAME")
    oneDrivePath = "C:\Users\" & userName & Worksheets("Info").Cells(3, 15).Value
    
    Dim folder As Object
    ' FileSystemObjectのインスタンスを作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(oneDrivePath)
    
    'パスにhttps:が含まれる場合
    If url Like "https:*" Then
        For Each subfolder In folder.SubFolders
            Dim subfolderName As String
            subfolderName = subfolder.name

            Dim subfolderPosition As Integer
            subfolderPosition = InStr(1, url, subfolderName & "/") ' subfolderNameの位置を取得

            If subfolderPosition > 0 Then
                Dim relativePath As String
                relativePath = Mid(url, subfolderPosition) ' subfolderNameを含む部分からのパスを取得
                GetPath = oneDrivePath & relativePath ' OneDriveパスにその部分を追加
                Exit Function
            End If
        Next subfolder
    Else
        GetPath = url 'httpsを含まない場合はそのまま返す
    End If
End Function
