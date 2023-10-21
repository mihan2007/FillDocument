Attribute VB_Name = "DeleteCompany"
Sub DeleteSelectedCompany()
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load MainXMLFilePath
    
    Dim comboBox As Object
    Set comboBox = LegalEntityProfile.Controls("ComboBox1")
    
    ' �������� ��������� ��� �� ComboBox1
    Dim selectedCompany As String
    selectedCompany = comboBox.Value
    
    ' ������� ��������������� ���� LegalEntity � ��������� CompanyName
    Dim nodes As Object
    Set nodes = xmlDoc.SelectNodes("//LegalEntity[@CompanyName='" & selectedCompany & "']")
    
    If nodes.Length > 0 Then
        ' ���� ������, ������� ���� LegalEntity
        xmlDoc.DocumentElement.RemoveChild nodes(0)
        
        ' ��������� ��������� � XML-�����
        xmlDoc.Save MainXMLFilePath
        
        ' �������� ComboBox1 � �������� ��� ��������
        comboBox.Clear
        Dim updatedNodes As Object
        Set updatedNodes = xmlDoc.SelectNodes("//LegalEntity/@CompanyName")
        For Each Node In updatedNodes
            comboBox.AddItem Node.Text
        Next Node
    Else
        MsgBox "�������� �� ������� � XML."
    End If
End Sub

