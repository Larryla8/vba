Attribute VB_Name = "XmlReaderModule"
Option Explicit

Sub ReadSample()

    ' �����ե�����
    Dim inputFileName As String
    inputFileName = "C:\work\scubism\ana0331.xml"
   
    ' XML���i���z�ߜʂ���Ф�
    Dim xr As XmlReader
    Set xr = New XmlReader
    xr.LoadXmlFile (inputFileName)
   
    ' XML���ǩ`�����i���z��
    Dim memberList() As XmlTypes.Member
    Call xr.GetMemberList(memberList)
   
    ' ȡ�ýY���򥻥�˳�������
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
