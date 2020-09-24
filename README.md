<div align="center">

## SIMPLE SQL ACCESS


</div>

### Description

This is Easy. Very Easy. 10 lines of code, 30 of comments. I'm sick to death of trying to find simple, concise code that tells me what I need to know. This is an example of how to access a SQL Database from VB, and how to use the data returned.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Zen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zen.md)
**Level**          |Beginner
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zen-simple-sql-access__1-11093/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub cmdConnect_Click()
  ' We're going to assume we're using a SQL Server
  ' that is named 'Server'. We're also going to assume
  ' the target database's name is 'Database'. The default
  ' UID in SQL is 'sa', so let's just use that, and the
  ' default password is either 'sa' or blank (I'm going to
  ' use blank). Finally, let's assume the Table name
  ' is 'Table'.
  Dim objConnection As Object
  Dim objContents As Object
  Dim strSQL As String
  ' Create the ADO connection. This is the handiest way to
  ' connect to a database in my uneducated opinion, so if you
  ' disagree, write your own code. ;-)
    Set objConnection = CreateObject("ADODB.Connection")
  ' Next, open the connection to the database.
    objConnection.Open "Driver={SQL Server};Server=Server;Database=Database;uid=sa;pwd=;"
  ' Now, for this next part to make sense, you'll need at least
  ' a little experience writing SQL queries. This is the simplest.
    strSQL = "SELECT * FROM Table"
  ' Finally, Create a Recordset using the SQL string we wrote above.
  ' What's happening here is the connection object (objConnection)
  ' is executing the SQL query, then building a recordset called
  ' objContents with the results returned from our query.
    Set objContents = objConnection.execute(strSQL)
  ' Lastly, I bet you're wondering how to get at that data. Well,
  ' If you're only interested in the first value returned, I
  ' recommend this, quick and easy.
    varResult = objContents(0)
  ' If you're looking to gather a value for a particular field
  ' in the table, this is the way to go. Just replace <FIELD NAME>
  ' with your field's name (you DO need the quotes).
    varResult = objContents("<FIELD NAME>")
  ' So if you wanted to return every value, you can simply use a
  ' while loop and BOF (Beginning Of File) and EOF (End Of File);
  ' SQL gets pissy if you try to go past the end of the file.
    While objContents.BOF = False And objContents.EOF = False
      varResult = objContents("<FIELD NAME>")
      ListBox1.AddItem varResult
 	  objContents.MoveNext ' This moves on to the next ROW
    Wend
End Sub
```

