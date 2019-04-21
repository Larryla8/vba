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

' コンストラクタ
Public Sub Class_Initialize()
End Sub

' XMLをDOMオブジェクトにロードする
Public Sub LoadXmlFile(ByVal fileName As String)
    ' MSXMLオブジェクトを生成
    Set xmlDocument = Nothing
    Set xmlDocument = New MSXML2.DOMDocument60
    xmlDocument.Load (fileName)
End Sub

' メンバリストを取得する
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
    
   
    ' XMLのmemberノードを取得する
    Set membersNode = xmlDocument.SelectSingleNode("//members")
    Dim i As Integer
    i = 0
    For Each memberNode In membersNode.ChildNodes
        ReDim Preserve memberList(i)
       
        ' idの属性値を取得する
        For Each memberAttribute In memberNode.Attributes
            If memberAttribute.name = "id" Then
                memberList(i).id = memberAttribute.Value
            End If
        Next memberAttribute
       
        ' memberの子要素を取得する
        Dim childNode As IXMLDOMNode
        For Each childNode In memberNode.ChildNodes
            ' name要素の値を取得する
            If childNode.nodeName = "name" Then
                memberList(i).name = childNode.Text
            End If
            ' age要素の値を取得する
            If childNode.nodeName = "age" Then
                memberList(i).age = childNode.Text
            End If
        Next childNode
       
        i = i + 1
    Next memberNode
   
End Function

' デストラクタ
Public Sub Class_Terminate()
    If Not xmlDocument Is Nothing Then Set xmlDocument = Nothing
End Sub
