Option Explicit

' ThisWorkbook.Sheets("VBA").TextBox8 : OpenAI_URL
' ThisWorkbook.Sheets("VBA").TextBox1 : OpenAI_API_Key
' ThisWorkbook.Sheets("VBA").ComboBox1 : OpenAI_Model
' ThisWorkbook.Sheets("VBA").CommandButton1 : <Submit request>
' ThisWorkbook.Sheets("VBA").CommandButton2 : <Check URL, API Key and Model>
' ThisWorkbook.Sheets("VBA").TextBox13 : OpenAI_Role

Private Sub CommandButton1_Click()

    ThisWorkbook.Sheets("VBA").TextBox3 = ""
    
    If (ThisWorkbook.Sheets("VBA").TextBox1 = "") Then
        'OpenAI API KEY from Environment variable
        ThisWorkbook.Sheets("VBA").TextBox1 = Environ("OPENAI_API_KEY")
    End If
    
    Call Processing_ChatGPT
    
End Sub
Private Sub CommandButton2_Click()

    If (ThisWorkbook.Sheets("VBA").TextBox1 = "") Then
        'OpenAI API KEY from Environment variable
        ThisWorkbook.Sheets("VBA").TextBox1 = Environ("OPENAI_API_KEY")
    End If


    ThisWorkbook.Sheets("VBA").TextBox9 = ""
    ThisWorkbook.Sheets("VBA").TextBox10 = ""
    ThisWorkbook.Sheets("VBA").TextBox11 = ""
    
    Call Test_URL_ChatGPT

End Sub
Private Sub TextBox13_Change()

    Dim strSystemRole As String
    Dim strSystemRoleJson As String
    
    strSystemRole = ThisWorkbook.Sheets("VBA").TextBox13
    
    strSystemRoleJson = OpenAI_InputRole2JSON("system", strSystemRole)
     
    ThisWorkbook.Sheets("VBA").TextBox14 = strSystemRoleJson
    
    Call TextBox2_Change
    
End Sub
Private Sub TextBox2_Change()

    Dim strInput As String
    Dim question As String
    Dim MODEL As String
    Dim strOpenAI_Model As String
    
    strOpenAI_Model = ThisWorkbook.Sheets("VBA").ComboBox1
    strInput = ThisWorkbook.Sheets("VBA").TextBox2

    
    question = OpenAI_InputRole2JSON("user", strInput)
    MODEL = OpenAI_Model2JSON(strInput, strOpenAI_Model)
    
    ThisWorkbook.Sheets("VBA").TextBox6 = question
    ThisWorkbook.Sheets("VBA").TextBox7 = MODEL
    
End Sub
