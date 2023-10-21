Attribute VB_Name = "DeleteCompany"
Sub DeleteSelectedCompany()
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load MainXMLFilePath
    
    Dim comboBox As Object
    Set comboBox = LegalEntityProfile.Controls("ComboBox1")
    
    ' Получите выбранное имя из ComboBox1
    Dim selectedCompany As String
    selectedCompany = comboBox.Value
    
    ' Найдите соответствующий узел LegalEntity с атрибутом CompanyName
    Dim nodes As Object
    Set nodes = xmlDoc.SelectNodes("//LegalEntity[@CompanyName='" & selectedCompany & "']")
    
    If nodes.Length > 0 Then
        ' Если найден, удалите узел LegalEntity
        xmlDoc.DocumentElement.RemoveChild nodes(0)
        
        ' Сохраните изменения в XML-файле
        xmlDoc.Save MainXMLFilePath
        
        ' Очистите ComboBox1 и обновите его значения
        comboBox.Clear
        Dim updatedNodes As Object
        Set updatedNodes = xmlDoc.SelectNodes("//LegalEntity/@CompanyName")
        For Each Node In updatedNodes
            comboBox.AddItem Node.Text
        Next Node
    Else
        MsgBox "Компания не найдена в XML."
    End If
End Sub

