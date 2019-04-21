Attribute VB_Name = "XmlReaderModule"
Option Explicit

Sub ReadSample()

    ' 入力ファイル
    Dim inputFileName As String
    inputFileName = "C:\work\scubism\ana0331.xml"
   
    ' XMLの読み込み準備を行う
    Dim xr As XmlReader
    Set xr = New XmlReader
    xr.LoadXmlFile (inputFileName)
   
    ' XMLよりデータを読み込む
    Dim memberList() As XmlTypes.Member
    Call xr.GetMemberList(memberList)
   
    ' 取得結果をセルに出力する
    If Sgn(memberList) <> 0 Then
   
        Dim rowIndex As Long
        rowIndex = 2
       
        Dim i As Integer
        For i = 0 To UBound(memberList)
            Cells(rowIndex, 1) = memberList(i).id
            Cells(rowIndex, 2) = memberList(i).name
            Cells(rowIndex, 3) = memberList(i).age
            rowIndex = rowIndex + 1
        Next i
    End If
   
    Set xr = Nothing
End Sub
