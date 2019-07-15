# CreateNewAccountIteration_VBA
This method (written in Visual Basic) creates a new account iteration. This is very helpful when creating several accounts at the same time, that follows a certain naming convention.

## Background Info
In my field of work, I am required to create multiple accounts per day throughout multiple web environments. These accounts have to follow a certain naming convention per environment.
A naming convention for a new account could be:

"node" + last name + environment suffix + #. 

For example, in our Test environment, "nodemccoytest200" would be a proper naming convention.

These accounts had to be accounted for, followed by the passwords, the default languages, and the country of origin. I created an Access database for this reason, as I enjoy the accessibility of forms and tabs.

Now, imagine having to type out essentially the same username, with account details, MULTIPLE. TIMES. A DAY.

I did this, for about 2 months... then I figured I didn't enjoy the feeling of carpal tunnel, so I created a VB script that did the dirty work for me.

## Take a look...

<img src="/Images/AcctImg01.jpg"/>

This is a screen capture of a form that I created to enter and edit all accounts for a Test environment. As you can see, I have created several accounts...

My latest account created is listed as "nodemccoytest280".

<img src="/Images/AcctImg02.jpg"/>

Here, when we select the "New Account" button at the bottom, a new account is created that follows the naming convention, and goes up one interval; "nodemccoytest281".

## The code...

<pre><code>
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
            
</code></pre>
At the start of our function, we need to declare our database and recordset variables in order to perform our OpenRecordset function.
The SELECT statement grabs the latest created record from it's ID (accID) from the table (tbl_Accounts), where its account type (accType) is equal to 'TST' (which is our suffix for our Test environment). It will then be initialized as the string variable 'strFindID'.

'strFindID' will then be passed into our OpenRecordset function and read. It will find the maxID, and then be initialized as 'varMaxID'.

<pre><code>
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
</code></pre>
Here we are at our second SELECT statement initialized as a string variable, 'strFindLastUser'. This time, it will SELECT the account name (accName) from the table (tbl_Accounts), where the ID (accID) is equal to our maximum ID that the compiler selected from the first SELECT statement (varMaxID).

The string will then be used in passing itself to our OpenRecordset function, and initialized as 'varMaxUser'. This time, 'varMaxUser' will be cleaned of any characters, that way only a number is returned.

This number will then be converted into an INT by the use of 'CInt', and then adding 1 to itself.

<pre><code>
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
</code></pre>
Here, we are initializing all of our variables to be read into the SELECT statement. This is so the SELECT statement is readable. This is also the time to change any variables you want by default, since every time the function is called, these variables will be inserted into the database.

<pre><code>
'-----------------String: Insert new Username into table--------------------------------------
Dim strInsert As String
strInsert = "INSERT INTO tbl_Accounts " & _
            "(accName, accEmail, accPassword, accLang, accCountry, accType) VALUES " & _
            "('" & finalUsername & "' , '" & finalEmail & "', '" & finalPassword & "', '" & finalLanguage & "' , " & CStr(finalCountry) & " ,'" & finalType & "');"
'---------------------------------------------------------------------------------------------

dbs.Execute (strInsert)
DoCmd.Requery

End Sub
</code></pre>
Here we are at our final form. The INSERT statement is initialzed as a string, and is used in our 'dbs.Execute' function, which executes the statement.

The form is then requeried to display the newly created record.

