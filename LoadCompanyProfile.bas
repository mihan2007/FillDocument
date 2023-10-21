Attribute VB_Name = "LoadCompanyProfile"
Public Const MainXMLFilePath As String = "C:\Tenders\CompanyListInfo.xml"
Public Const ChoisenCompanyXMLFilePath As String = "C:\Tenders\DefaultCompany.xml"

Sub LoadXMLFeileToUserForm()
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load MainXMLFilePath
    
    Dim comboBox As Object
    Set comboBox = LegalEntityProfile.Controls("ComboBox1")
    comboBox.Clear
    
    Dim nodes As Object
    Set nodes = xmlDoc.SelectNodes("//LegalEntity")
 
    For Each Node In nodes
        Dim companyName As String
        companyName = Node.getAttribute("CompanyName")
        comboBox.AddItem companyName
    Next Node
End Sub

Public Sub LoadChoisenCompanyLable()

    Dim choisenXMLDoc As Object
    Set choisenXMLDoc = CreateObject("MSXML2.DOMDocument.6.0")
    choisenXMLDoc.async = False
    choisenXMLDoc.Load ChoisenCompanyXMLFilePath
    
    Dim ChoisenCompanyName As String
 ChoisenCompanyName = choisenXMLDoc.SelectSingleNode("//LegalEntity").getAttribute("CompanyName")
    
    LegalEntityProfile.Controls("ChoisenCompanyName").Caption = "Выбрано: " & ChoisenCompanyName
    
End Sub

Sub UpdateLegalEntityProfile(selectedName As String)


    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    
    If Dir(MainXMLFilePath) = "" Then
        MsgBox "The XML file does not exist.", vbExclamation
        Exit Sub
    End If
    
    xmlDoc.Load MainXMLFilePath
    
    Dim selectedNode As Object
    Set selectedNode = xmlDoc.SelectSingleNode("//LegalEntity[@CompanyName='" & selectedName & "']")
    
    If Not selectedNode Is Nothing Then
        Dim fieldNames() As Variant
        fieldNames = Array("Address", "PhoneNumber", "Email", "INN", "KPP", "OGRN", "DateOfBirth", "OKVED", "GeneralManager", "Passport", "Email", "AccountDetail")
        
        Dim i As Integer
        For i = LBound(fieldNames) To UBound(fieldNames)
            Dim fieldText As Variant
            On Error Resume Next
            Set fieldText = selectedNode.SelectSingleNode(fieldNames(i))
            On Error GoTo 0
            
            If Not fieldText Is Nothing Then
                LegalEntityProfile.Controls("TextBox" & fieldNames(i)).Text = fieldText.Text
            End If
        Next i
        LegalEntityProfile.TextBoxCompanyName.Value = selectedName
        LegalEntityProfile.Controls("ComboBox1").Value = selectedName
    End If
    
    DisableTextBoxes
End Sub



Sub DisableTextBoxes()
    Dim ctrl As Control

    For Each ctrl In LegalEntityProfile.Controls

        If TypeOf ctrl Is MSForms.TextBox Then

            ctrl.Locked = True
        End If
    Next ctrl
End Sub

Sub EnableTextBoxes()
    Dim ctrl As Control

    For Each ctrl In LegalEntityProfile.Controls

        If TypeOf ctrl Is MSForms.TextBox Then

            ctrl.Locked = False
        End If
    Next ctrl
End Sub

