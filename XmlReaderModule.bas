Attribute VB_Name = "XmlReaderModule"
Option Explicit

Sub ReadXML()

    Dim inputFileName As String
    
    ' �t�@�C�����擾
    inputFileName = Sheets("Sheet1").Range("C1")

    If inputFileName = "" Then
        MsgBox "�t�@�C���p�X�ƃt�@�C��������͉������B"
        Exit Sub
    
    End If
    
   
    'XML�t�@�C���ǂݍ���
    Dim xr As XmlReader
    Set xr = New XmlReader
    xr.LoadXmlFile (inputFileName)
   
    'XML����K�v�ȏ��擾
     Call xr.GetMemberList
   
   
    Set xr = Nothing
End Sub
