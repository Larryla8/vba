VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DOM
Private xmlDocument As MSXML2.DOMDocument60

' ���󥹥ȥ饯��
Public Sub Class_Initialize()
End Sub

' XML��DOM���֥������Ȥ˥��`�ɤ���
Public Sub LoadXmlFile(ByVal fileName As String)
    ' MSXML���֥������Ȥ�����
    Set xmlDocument = Nothing
    Set xmlDocument = New MSXML2.DOMDocument60
    xmlDocument.Load (fileName)
End Sub

' ���Хꥹ�Ȥ�ȡ�ä���
Public Function GetMemberList(ByRef memberList() As XmlTypes.Member)
    Dim membersNode As IXMLDOMNode
    Dim memberNode As IXMLDOMNode
    Dim memberAttribute As MSXML2.IXMLDOMAttribute
    
    Dim filesNode As IXMLDOMNode
    
    Dim fileNode As IXMLDOMNode
    
    Dim issusList As IXMLDOMNodeList
    
    Dim issuList As IXMLDOMNodeList
    Dim ruleList As IXMLDOMNodeList
    Dim codeList As IXMLDOMNodeList
   
    Dim cntIssues As Integer
    Dim cntCode As Integer
    

    Set filesNode = xmlDocument.SelectSingleNode("//project-issues/files")
    
    
    For Each fileNode In filesNode.ChildNodes
        
        Set issusList = fileNode.ChildNodes.Item(2).ChildNodes
        

        
        For cntIssues = 0 To issusList.Length - 1
            Set issuList = issusList.Item(cntIssues).ChildNodes
            
            Set ruleList = issuList.Item(1).ChildNodes
            
            Set codeList = issuList.Item(3).ChildNodes
            
            If ruleList.Item(4).Text = "high" Then
                Debug.Print vbLf
                Debug.Print fileNode.ChildNodes.Item(0).Text
                Debug.Print "-------------------------------- -"
                Debug.Print "Status: " & issuList.Item(2).Text
                
                Debug.Print "Importance: " & ruleList.Item(4).Text
                Debug.Print "Rule Category : " & ruleList.Item(2).Text
                Debug.Print "Rule Title : " & ruleList.Item(3).Text
                Debug.Print "Code Snippet: " & vbLf
                For cntCode = 0 To codeList.Length - 1
                    Debug.Print codeList.Item(cntCode).Attributes(0).Text; ": " & codeList.Item(cntCode).Text
                Next
                
            End If
                
        Next
                
 
    
    
    
    
    Next
    
   
    ' XML��member�Ω`�ɤ�ȡ�ä���
    Set membersNode = xmlDocument.SelectSingleNode("//members")
    Dim i As Integer
    i = 0
    For Each memberNode In membersNode.ChildNodes
        ReDim Preserve memberList(i)
       
        ' id�����Ԃ���ȡ�ä���
        For Each memberAttribute In memberNode.Attributes
            If memberAttribute.name = "id" Then
                memberList(i).id = memberAttribute.Value
            End If
        Next memberAttribute
       
        ' member����Ҫ�ؤ�ȡ�ä���
        Dim childNode As IXMLDOMNode
        For Each childNode In memberNode.ChildNodes
            ' nameҪ�ؤ΂���ȡ�ä���
            If childNode.nodeName = "name" Then
                memberList(i).name = childNode.Text
            End If
            ' ageҪ�ؤ΂���ȡ�ä���
            If childNode.nodeName = "age" Then
                memberList(i).age = childNode.Text
            End If
        Next childNode
       
        i = i + 1
    Next memberNode
   
End Function

' �ǥ��ȥ饯��
Public Sub Class_Terminate()
    If Not xmlDocument Is Nothing Then Set xmlDocument = Nothing
End Sub