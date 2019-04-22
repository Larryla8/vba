Attribute VB_Name = "XmlReaderModule"
Option Explicit

Sub ReadXML()

    Dim inputFileName As String
    
    ' ファイルを取得
    inputFileName = Sheets("Sheet1").Range("C1")

    If inputFileName = "" Then
        MsgBox "ファイルパスとファイル名を入力下さい。"
        Exit Sub
    
    End If
    
   
    'XMLファイル読み込み
    Dim xr As XmlReader
    Set xr = New XmlReader
    xr.LoadXmlFile (inputFileName)
   
    'XMLから必要な情報取得
     Call xr.GetMemberList
   
   
    Set xr = Nothing
End Sub
