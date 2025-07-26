' ThisWorkbook.Sheets("VBA").CommandButton1 : <Submit request>
' ThisWorkbook.Sheets("VBA").CommandButton2 : <Check URL, API Key and Model>
' ThisWorkbook.Sheets("VBA").ComboBox1 : OpenAI_Model
' ThisWorkbook.Sheets("VBA").TextBox1 : OpenAI_API_Key
' ThisWorkbook.Sheets("VBA").TextBox2 : Input_Text
' ThisWorkbook.Sheets("VBA").TextBox3 : Output_Text
' ThisWorkbook.Sheets("VBA").TextBox4 : Output_Text_JSON
' ThisWorkbook.Sheets("VBA").TextBox5 : Request_OpenAI_Status
' ThisWorkbook.Sheets("VBA").TextBox6 : Input_Text_JSON
' ThisWorkbook.Sheets("VBA").TextBox7 : Request
' ThisWorkbook.Sheets("VBA").TextBox8 : OpenAI_URL
' ThisWorkbook.Sheets("VBA").TextBox9 : Internet_Status
' ThisWorkbook.Sheets("VBA").TextBox10 : Details_Model
' ThisWorkbook.Sheets("VBA").TextBox11 : Request_Status
' ThisWorkbook.Sheets("VBA").TextBox13 : OpenAI_SystemRole
' ThisWorkbook.Sheets("VBA").TextBox14 : OpenAI_SystemRole_JSON
' ThisWorkbook.Sheets("VBA").TextBox15 : Number_Tokens
' ThisWorkbook.Sheets("VBA").TextBox16 : Price

Private Sub Workbook_Open()

     Dim sourceRange As Excel.Range
    Dim cb As ComboBox
    Dim strInput As String
    Dim question As String
    Dim request As String
    Dim strOpenAI_Model As String

    Set WS = ThisWorkbook.Sheets("VBA")
    
    Set cb = ThisWorkbook.Sheets("VBA").ComboBox1
    Set sourceRange = ThisWorkbook.Sheets("params").Range("A2:A33")
    cb.List = sourceRange.Value
    
    ThisWorkbook.Sheets("VBA").ComboBox1 = "gpt-4.1"
    
    'Other fields
    ThisWorkbook.Sheets("VBA").TextBox1 = ""
    ThisWorkbook.Sheets("VBA").TextBox2 = ""
    ThisWorkbook.Sheets("VBA").TextBox3 = ""
    ThisWorkbook.Sheets("VBA").TextBox4 = ""
    ThisWorkbook.Sheets("VBA").TextBox5 = ""
    ThisWorkbook.Sheets("VBA").TextBox6 = ""
    ThisWorkbook.Sheets("VBA").TextBox7 = ""
    ThisWorkbook.Sheets("VBA").TextBox9 = ""
    ThisWorkbook.Sheets("VBA").TextBox10 = ""
    ThisWorkbook.Sheets("VBA").TextBox11 = ""
    
    strOpenAI_Model = ThisWorkbook.Sheets("VBA").ComboBox1
    strInput = ThisWorkbook.Sheets("VBA").TextBox2
    
    question = OpenAI_InputRole2JSON("user", strInput)
    request = OpenAI_Model2JSON(strInput, strOpenAI_Model)
    
    ThisWorkbook.Sheets("VBA").TextBox6 = question
    ThisWorkbook.Sheets("VBA").TextBox7 = request

End Sub
