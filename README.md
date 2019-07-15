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
            
</code></pre>
