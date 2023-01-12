# Excel-VBA-SQL
This code shows the fastest way to insert data from an excel sheet to an online server using SQL Server and VBA


Option Explicit
Dim cn As New ADODB.Connection
Dim i, j, LastRow, vbox As Integer
Dim SQL_command, a, b, c, var1, var2, vpo, var3, var4, var5, var6, id As String
Dim ws As Worksheet

Sub insert_from_EXCEL_to_DB()
        
    Set ws = Sheets("extract")
    LastRow = Range("A1048576").End(xlUp).Row

#creating a connection with the server

    cn.Open "Provider=SQLOLEDB.1;Server=SERVER_URL;database=DATABASE_NAME;Database=DATABASE_NAME;UID=USER_NAME;PWD=PASSWORD;"
   
#A temporary table has been created to support the merge job. Thus, on each job, the table is cleared.
    
    SQL_command = "DELETE FROM TEMP_TABLE;"
    cn.Execute SQL_command
    
#This is the fastest way to upload the information from excel to a SQL database, since we first define the second line(skipping the header) and then we connect to the rest of the table. So, only one query with all data is triggered.

    For i = 2 To 2
        a = ws.Range("A" & i)
        b = ws.Range("B" & i)
        id = a & b
        var1 = ws.Range("A" & i)
        var2 = ws.Range("B" & i)
        vpo = ws.Range("C" & i)
        var3 = ws.Range("D" & i)
        var4 = ws.Range("E" & i)
        var5 = ws.Range("F" & i)
        var6 = ws.Range("G" & i)
        vbox = ws.Range("K" & i)
    
#creating the first part of the query (only row 2)

        SQL_command = "INSERT TEMP_EXTRACT (VAR1, VAR2, EPO, VAR3, VAR4, VAR5, VAR6, EBOX, ID) VALUES ('" & var1 & "', '" & var2 & "', '" & vpo & "', '" & var3 & "', '" & var4 & "', '" & var5 & "', '" & var6 & "', '" & "', '" & vbox & "', '" & id & "')"
        
#creating the second part of the query (the left rows)       
        For j = 3 To LastRow
        a = ws.Range("A" & j)
        b = ws.Range("B" & j)
        id = a & b
        var1 = ws.Range("A" & j)
        var2 = ws.Range("B" & j)
        vpo = ws.Range("C" & j)
        var3 = ws.Range("D" & j)
        var4 = ws.Range("E" & j)
        var5 = ws.Range("F" & j)
        var6 = ws.Range("G" & j)
        vbox = ws.Range("K" & j)
        
        
        SQL_command = SQL_command & ",('" & var1 & "', '" & var2 & "', '" & vpo & "', '" & var3 & "', '" & var4 & "', '" & var5 & "', '" & var6 & "', '" & vbox & "', '" & id & "')"
    
    
        Next j

#Executing the entire query with only one command

        cn.Execute SQL_command

    Next i
    

#Merging the temporary table (cleaned and filled as above) with the destination Database    
SQL_command = "MERGE dbo.DEST_DB AS D " & _
      "USING dbo.TEMP_EXTRACT AS O ON (O.EPKEY = D.PKEY) " & _
      "WHEN MATCHED THEN " & _
      "UPDATE SET " & _
      "D.VAR6 = O.VAR6, D.STATUS_EXTRACTION = 'IN' " & _
      "WHEN NOT MATCHED BY SOURCE THEN " & _
      "UPDATE SET " & _
      "D.STATUS_EXTRACTION = 'NOT IN' " & _
      "WHEN NOT MATCHED BY TARGET THEN " & _
      "INSERT (VAR1, VAR2, EPO, VAR3, VAR4, VAR5, VAR6, EBOX, ID) " & _
      " VALUES (O.VAR1, O.VAR2, O.EPO, O.VAR3, O.VAR4, O.VAR5, O.VAR6, O.EBOX, O.ID,'IN');"

cn.Execute SQL_command

        MsgBox "Extraction Sucessfully."
        cn.Close
        Set cn = Nothing
