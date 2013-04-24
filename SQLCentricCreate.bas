Attribute VB_Name = "SQLCentricCreate"
Option Compare Database
' SQLCentricCreate
'
' This module has utility routines related to creating things in a more
' SQL centric (non GUI) manner.  This includes things like creating tables
' via SQL DDL, creating queries from (commented) SQL source, etc.
'
' See code in CreateTables() for SQL DDL and modify it to suit.
' See code in Xxx for creating queries
'

Sub CreateTables()
' ----------------------------------------------------------------------------
' Creates tables and relationships.  SQL DDL is embedded in the code.
' The prototypical flow here is:
'
'   add a comment to indicate waht this chunk of SQL is creating
'
'   start with an empty string (sql = "")
'
'   blather in the SQL DDL wrapped in VB code to keep appending to one long string
'
'   execute the sql, stripping comments at the same time
'
'   Optionally rifle through the table fields and add descriptions
'
' Currently there is no support for dropping, so if you want to re-run
' you need to delete the tables by hand.  Note that you need to hit F5
' to refresh the navigator to see the newly created tables.
'
' Look in database tools -> Relationships to see the ERD
'
' See comments in stripComments(...) for moving back and forth from VBA
' to just SQL.
'
' Here is an example call (suitable for running in the immediate window):
'
'     Call CreateTables()


    Dim sql As String
    Dim db As DAO.Database
    Dim strSQL As String
    
    Set db = CurrentDb
    
    ' -- Cars -----------------------------------------------------------------------
    sql = ""
    sql = sql & "-- Create the Cars table" & vbCrLf
    sql = sql & "CREATE TABLE Cars" & vbCrLf
    sql = sql & "   (" & vbCrLf
    sql = sql & "      CarID LONG," & vbCrLf
    sql = sql & "      CarName TEXT(50)," & vbCrLf
    sql = sql & "      ColorID LONG" & vbCrLf
    sql = sql & "   );" & vbCrLf
    
    ' Create it
    db.Execute stripComments(sql)
    
    ' Set the descriptions on the fields just created
    Call setFieldDescription("Cars", "CarID", "Primary Key")
    Call setFieldDescription("Cars", "CarName", "This is the model name of the car")
    Call setFieldDescription("Cars", "ColorID", "Foreign Key to Colors")
    ' -- End of Cars ------------------------------------------------------------------



    ' -- Colors -----------------------------------------------------------------------
    sql = ""
    sql = sql & "" & vbCrLf
    sql = sql & " -- Create the Colors table" & vbCrLf
    sql = sql & " CREATE TABLE Colors" & vbCrLf
    sql = sql & "   (" & vbCrLf
    sql = sql & "      ColorID LONG CONSTRAINT PK_Colors PRIMARY KEY," & vbCrLf
    sql = sql & "      ColorName TEXT(50)" & vbCrLf
    sql = sql & "   );" & vbCrLf
    sql = sql & "" & vbCrLf
    
    ' Create it
    db.Execute stripComments(sql)
    
    ' Set the descriptions on the fields just created
    Call setFieldDescription("Colors", "ColorID", "Primary Key")
    Call setFieldDescription("Colors", "ColorName", "This is the name of the color")
    ' -- End of Colors ------------------------------------------------------------------
    
    
    
    ' -- Foreign key from Cars to Colors -------------------------------------------------
    sql = ""
    sql = sql & "-- Add a foreign key constraint" & vbCrLf
    sql = sql & " ALTER TABLE Cars" & vbCrLf
    sql = sql & "   ADD CONSTRAINT MyColorIDRelationship" & vbCrLf
    sql = sql & "   FOREIGN KEY (ColorID) REFERENCES Colors (ColorID)" & vbCrLf
    sql = stripComments(sql)
    
    ' Create it
    db.Execute stripComments(sql)
    ' -- End of Foreign key from Cars to Colors --------------------------------------------

End Sub

Sub setFieldDescription(tableName As String, fieldName As String, propertyString As String)
' This routine sets a field's decription.  Parameters:
'   tableName - Name of the table the field is in (String)
'   fieldName - Name of the field (String)
'   propertyString - String you want property set to (String)
'
' Here is an example call (suitable for running in the immediate window):
'
'     Call setFieldDescription("Colors", "ColorID", "Enter Color ID")

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim prp As DAO.Property
    Dim PROPERTY_TYPE As String
    PROPERTY_TYPE = "Description"
    
    ' Get the field object for the table name and field name passed to us
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    Set fld = tdf.Fields(fieldName)

    ' If it exists, set it.  Otherwise get the 3270 error and create it.
    On Error Resume Next
    ' it exists, set it
    fld.Properties(PROPERTY_TYPE) = propertyString

    ' It's new, create it
    If Err.Number = 3270 Then
        Set prp = fld.CreateProperty(PROPERTY_TYPE, dbText, propertyString)
        fld.Properties.Append prp
        fld.Properties.Refresh
    End If

End Sub

' ------------------------------------------------------------------------------
' Simple comment stripper.  Takes a string, scans it for comments which are
' defined as text from "--" to the end of a line.  Removes the comments and
' blank lines and returns the modified string.
'
' If you're developping your query in a text editor, searching for this
' regex pattern and replacing with nothing will do the equivalent:
'
'    s/(^--.*\n)|--.*$| +--.*\n//

' While not strictly part of comment stripping, this is a nice central
' place to note the strategy for moving back and forth from VBA to SQL
' like you might use in SQL*Plus or the like.
' ----------------------------------------------------------------------------
' From text editor to VB code:
'
' 1) Double quote any existing quotes (i.e. s/"/""/)
' 2) Add    sql = sql & "    to the start of each line:
'
'        s/^/sql = sql & "/
'
'    Add    " & vbCrLf   to the end of each line:
'    NB - there should be a single leading space before the &
'
'        s/$/" & vbCrLf/
'
' From VB code to text editor
'
' 1) Reverse the above, lopping off what go added:
'
'        s/sql = sql & "//
'        s/" & vbCrLf//

Private Function stripComments(sText As String) As String
    ' Attempts to remove comments using a simple implementation
    ' Look for -- and strip from there to end of line

    Dim commentStart As Integer
    Dim endOfLine As Integer
    Dim sTmp As String
    
    ' init the loop by looking for the first "--" occurrence.  InStr returns 0 if not found
    ' or the position in the string
    
    commentStart = InStr(1, sText, "--")
    
    While commentStart > 0

        'If we're in the loop, we found the start of a comment.  Now find
        'the end of the line (CRLF) so we can strip the comment out
        endOfLine = InStr(commentStart, sText, vbCrLf)
        
        'If we didn't find an end of line (i.e. CRLF), then fallback to
        'just lopping off until the end of the srtring
        If (Not (endOfLine > 0)) Then endOfLine = Len(sText)

        ' Update the string to strip the current comment.  Grab the
        ' substring before the comment, and the string after the comment
         sText = Left(sText, commentStart - 1) & Mid(sText, endOfLine)
    
        ' update the loop condition
        commentStart = InStr(1, sText, "--")
    Wend

    ' After the stripping above, we may have blank lines.  These will be two CR LFs in a row
    ' In a loop replace each occurrence of two newlines in a row with just a single one.
    While (InStr(sText, vbCrLf & vbCrLf) > 0)
        sText = Replace(sText, vbCrLf & vbCrLf, vbCrLf)
    Wend
    
    ' Finally, there may be a blank line at the very beginning (just a single
    ' newline (CR LF) all by itself at the beginning).  Get rid of that too.
    If (Left(sText, 2) = vbCrLf) Then sText = Mid(sText, 3)
    
    ' Return the stripped string
    stripComments = sText

End Function

Sub TestIt()

    Call setFieldDescription("Colors", "ColorID", "Enter Color ID")

End Sub
