Attribute VB_Name = "ShowCompanyProfile"
Sub RunProgramm()
LegalEntityProfile.Show
End Sub
Sub CompanyNameWrite()
ReadAndWrite "CompanyName"
End Sub
Sub AddressWrite()
ReadAndWrite ("Address")
End Sub
Sub PhoneNumberWrite()
ReadAndWrite ("PhoneNumber")
End Sub
Sub EmailWrite()
ReadAndWrite ("Email")
End Sub
Sub INNWrite()
ReadAndWrite ("INN")
End Sub
Sub KPPWrite()
ReadAndWrite ("KPP")
End Sub
Sub OGRNWrite()
ReadAndWrite ("OGRN")
End Sub
Sub DateOfBirthWrite()
ReadAndWrite ("DateOfBirth")
End Sub
Sub OKVEDWrite()
ReadAndWrite ("OKVED")
End Sub
Sub GeneralManagerWrite()
ReadAndWrite ("GeneralManager")
End Sub
Sub PassportWrite()
ReadAndWrite ("Passport")
End Sub

Sub AccountDetailWrite()
ReadAndWrite ("AccountDetail")
End Sub

Sub EmailWirte()
ReadAndWrite ("Email")
End Sub


Function ReadAndWrite(fieldName As String)
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    ' Загрузка XML файла
    xml.async = False
    xml.Load ChoisenCompanyXMLFilePath
    
    ' Проверка на успешную загрузку XML
    If xml.parseError.ErrorCode <> 0 Then
        MsgBox "Ошибка при загрузке XML файла: " & xml.parseError.reason
        Exit Function
    End If
    
    ' Найти узел с именем поля
    Dim fieldNode As Object
    
    If fieldName = "CompanyName" Then
        Dim companyNameAttribute As Object
        Set companyNameAttribute = xml.SelectSingleNode("/Root/LegalEntity/@CompanyName")
        Selection.TypeText companyNameAttribute.Text
        Exit Function
    Else
        Set fieldNode = xml.SelectSingleNode("/Root/LegalEntity/" & fieldName)
    End If
    
    ' Проверка на наличие узла с именем поля
    If Not fieldNode Is Nothing Then
        ' Вывести значение узла в текущую позицию
        Selection.TypeText fieldNode.Text
    Else
        MsgBox "Узел <" & fieldName & "> не найден в XML файле."
    End If
End Function


