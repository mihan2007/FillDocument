Attribute VB_Name = "SaveCompanyProfile"
Sub SaveDataToXML(useMainXMLLogic As Boolean, updateExistingRecord As Boolean)
    
    Dim selectedName As String
    Dim xmlDoc As Object
    Dim fieldNames() As Variant
    Dim selectedNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    
    Dim updatedCompanyName As String
    updatedCompanyName = LegalEntityProfile.Controls("TextBoxCompanyName").Text
        
    
    selectedName = LegalEntityProfile.Controls("ComboBox1").Text
    
    fieldNames = Array("CompanyName", "Address", "PhoneNumber", "Email", "INN", "KPP", "OGRN", "DateOfBirth", "OKVED", "GeneralManager", "Passport", "Email", "AccountDetail")
    
    If useMainXMLLogic Then
        If Dir(MainXMLFilePath) = "" Then
            MsgBox "The XML file does not exist. Cannot save data.", vbExclamation
            Exit Sub
        End If
        xmlDoc.Load MainXMLFilePath
    Else
         
        xmlDoc.appendChild xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
        Set rootElement = xmlDoc.createElement("Root")
        xmlDoc.appendChild rootElement
    End If
    
    
If updateExistingRecord Then
    
    Set selectedNode = xmlDoc.SelectSingleNode("//LegalEntity[@CompanyName='" & selectedName & "']")
    
    If selectedNode Is Nothing Then
        MsgBox "The selected record does not exist. Cannot update data.", vbExclamation
        Exit Sub
    Else


        Dim originalCompanyName As String
        originalCompanyName = selectedNode.getAttribute("CompanyName")
        
        For Each fieldName In fieldNames
            Dim fieldValue As String
            fieldValue = LegalEntityProfile.Controls("TextBox" & fieldName).Text
            
            If fieldValue <> "" Then
                If fieldName <> "CompanyName" Then
                    Set fieldElement = selectedNode.SelectSingleNode(fieldName)
                    If Not fieldElement Is Nothing Then
                        fieldElement.Text = fieldValue
                    End If
                End If
            End If
        Next fieldName
        
        selectedNode.setAttribute "CompanyName", updatedCompanyName
    
        End If

    Else
        Set newLegalEntityNode = xmlDoc.createElement("LegalEntity")
        newLegalEntityNode.setAttribute "CompanyName", updatedCompanyName
        
        For Each fieldName In fieldNames
            'Dim fieldValue As String
            fieldValue = LegalEntityProfile.Controls("TextBox" & fieldName).Text
            
            ' Если поле пустое, устанавливаем "N/A"
            If fieldValue = "" Then
                fieldValue = "N/A"
            End If
                If fieldName <> "CompanyName" Then
                    Set fieldElement = xmlDoc.createElement(fieldName)
                    fieldElement.Text = fieldValue
                    newLegalEntityNode.appendChild fieldElement
                End If
        Next fieldName
        
        xmlDoc.DocumentElement.appendChild newLegalEntityNode
          
    End If
    
    If useMainXMLLogic Then
        xmlDoc.Save MainXMLFilePath
        MsgBox "Данные " & selectedName & " были обновленны " & MainXMLFilePath & ".", vbInformation
        UpdateLegalEntityProfile updatedCompanyName
    Else
        xmlDoc.Save ChoisenCompanyXMLFilePath
        MsgBox "Данные " & selectedName & " сделанны основными " & ChoisenCompanyXMLFilePath & ".", vbInformation
    End If
End Sub

