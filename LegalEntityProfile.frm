VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LegalEntityProfile 
   Caption         =   "Редактор Профилей"
   ClientHeight    =   8550.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9315.001
   OleObjectBlob   =   "LegalEntityProfile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LegalEntityProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public AddCompanyMode As Boolean
Public EditMode As Boolean

Private Sub AddCompanyButton_Change()

    If AddCompanyButton.Value Then

                ClearWindow
                EnableTextBoxes
                ComboBox1.Enabled = False
                EditButton.Enabled = False
                DeleteCompany.Enabled = False
                SelectDefaultCompanyButton.Enabled = False

    Else
        question = QuestionWindow("Добавление организации", "Вы уверены, что хотите создать организацию")
                If question = True Then
                
                SaveDataToXML True, False
                LoadXMLFeileToUserForm
                ComboBox1.Enabled = True
                EditButton.Enabled = True
                DeleteCompany.Enabled = True
                SelectDefaultCompanyButton.Enabled = True
                
                Else
                
                ComboBox1.Enabled = True
                EditButton.Enabled = True
                DeleteCompany.Enabled = True
                SelectDefaultCompanyButton.Enabled = True
                ChoseLastItemInCombobox
                End If
        
    End If
End Sub

Sub ClearWindow()
    Dim ctrl As Control
    For Each ctrl In LegalEntityProfile.Controls

        If TypeOf ctrl Is MSForms.TextBox Then

            ctrl.Value = ""
        End If
    Next ctrl
    
    LegalEntityProfile.ComboBox1.ListIndex = -1

End Sub


Private Sub ComboBox1_Change()
    Dim selectedName As String
    selectedName = LegalEntityProfile.Controls("ComboBox1").Text
 'ClearWindow
    UpdateLegalEntityProfile selectedName
    
    EditMode = False
    
End Sub

Private Sub DeleteCompany_Click()

    AreYouSureQuestion = QuestionWindow("Удаление Организации", "Вы уверенны, что хотите удалить организацию")
        
        If AreYouSureQuestion = True Then
            DeleteSelectedCompany
            ClearWindow
        Else
        
        End If
        
End Sub

Private Sub EditButton_Change()
    If AddCompanyMode = False Then
        If EditButton.Value Then
            EnableTextBoxes
            LegalEntityProfile.EditButton.Caption = "Сохранить"
            ComboBox1.Enabled = False
            AddCompanyButton.Enabled = False
            DeleteCompany.Enabled = False
            SelectDefaultCompanyButton.Enabled = False
            
        Else
    
            DisableTextBoxes
    
            SaveDataToXML True, True
            'SaveExistingDataToXML
                        
            Dim selectedName As String
            selectedName = LegalEntityProfile.Controls("ComboBox1").Text
            LoadXMLFeileToUserForm
            UpdateLegalEntityProfile selectedName
        
            LegalEntityProfile.EditButton.Caption = "Редактировать"
            
            ComboBox1.Enabled = True
            AddCompanyButton.Enabled = True
            DeleteCompany.Enabled = True
            SelectDefaultCompanyButton.Enabled = True
            
        End If
    Else
    EditButton.Enabled = True
    EditButton.Value = False
    End If
End Sub

Private Sub SelectDefaultCompanyButton_Click()

    If AddCompanyMode = False Then
    SaveDataToXML False, False
        LegalEntityProfile.ChoisenCompanyName.Caption = ""
        LoadChoisenCompanyLable
    Else
    End If
     
End Sub

Function QuestionWindow(ByVal title As String, ByVal questionText As String) As Boolean

    Dim response As VbMsgBoxResult
    
    response = MsgBox(questionText, vbYesNo + vbQuestion, title)

    ShowYesNoMessageBox = (response = vbYes)
    
    If response = vbYes Then
        QuestionWindow = True
    Else
        QuestionWindow = False
    End If
End Function

Sub ChoseLastItemInCombobox()
    If ComboBox1.ListCount >= 0 Then
    ComboBox1.ListIndex = 0
    Else

    End If
End Sub

Public Sub UserForm_Initialize()

    LoadXMLFeileToUserForm
    
    LoadChoisenCompanyLable
    
    If LegalEntityProfile.Visible = False Then

    End If
    
    AddCompanyMode = False
    
    ChoseLastItemInCombobox
    
End Sub
