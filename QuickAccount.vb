Private Sub cmdQuickAccountTest_Click()

'----------------Declarations----------------------

Dim dbs As DAO.Database, rst As DAO.Recordset, rst2 As DAO.Recordset
Dim varMaxID As Integer
Dim varMaxUser As String
Dim varUserNumber As String


'----------------String: Latest ID-------------------
Dim strFindID As String
strFindID = "SELECT Max(accID) as maxID " _
            & "FROM tbl_Accounts " _
            & "WHERE accType = 'TST';"
                       

'----------------Method: Get latest ID --------------------
Set dbs = CurrentDb
Set rst = dbs.OpenRecordset(strFindID)
varMaxID = rst!maxID


'----------------String: Latest username (from ID) -------------------
Dim strFindLastUser As String
strFindLastUser = "SELECT accName " & _
                "FROM tbl_Accounts " & _
                "WHERE accID = " & varMaxID & ""
                

'----------------Method: Get latest Username, then clean string to find the number-----------

Set rst2 = dbs.OpenRecordset(strFindLastUser)
varMaxUser = rst2!accName
varUserNumber = CleanString(varMaxUser)

Dim finalNumber As Integer
finalNumber = CInt(varUserNumber)
finalNumber = finalNumber + 1

'---------------------------------------------------------------------------------------------

'-----------------Method: Declare all new variables before insert-----------------------------
Dim finalUsername As String
Dim finalEmail As String
Dim finalPassword As String
Dim finalLanguage As String
Dim finalCountry As Integer
Dim finalType As String

finalUsername = "nodemccoytest" & finalNumber
finalEmail = "* +test" & finalNumber & "@gmail.com"
finalPassword = "Password123!"
finalLanguage = "English (EN)"
finalCountry = 233
finalType = "TST"

'-----------------String: Insert new Username into table--------------------------------------
Dim strInsert As String
strInsert = "INSERT INTO tbl_Accounts " & _
            "(accName, accEmail, accPassword, accLang, accCountry, accType) VALUES " & _
            "('" & finalUsername & "' , '" & finalEmail & "', '" & finalPassword & "', '" & finalLanguage & "' , " & CStr(finalCountry) & " ,'" & finalType & "');"
'---------------------------------------------------------------------------------------------

dbs.Execute (strInsert)
DoCmd.Requery


End Sub
