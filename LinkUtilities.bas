Attribute VB_Name = "LinkUtilities"
' This module has code to update link paths.
'
' UpdateLinks()
' Deals with various link types and sets everything relative
' to the current database directory
'
' UpdateCSVFolderLinks(Path As String)
' Just does links to text files and takes the path relative to the
' current database as a parameter.

Public Function UpdateLinks()

    ' update connections to linked databases in the same folder

    Dim tdf As dao.TableDef
    Dim ConnectString As String
    Dim ConnectStringFirstPart  As String
    Dim databaseNameStart As Integer
    Dim knownConnectionType As Boolean
    Dim fileName As String

    ' Loop through all the items in the TableDefs collection
    ' and update for specific types
    For Each tdf In CurrentDb.TableDefs

        knownConnectionType = False
        ConnectString = tdf.Connect

        ' This code has not been extensively tested for all
        ' connection/file types.  It works reliably for these known
        ' types: "Text" (i.e. CSV in this scenario), .accdb Acecess
        ' Databases, and Excel (Excel 8.0 xls). It may work for others
        ' too, that's just not been tested.

        If InStr(1, ConnectString, "Text", vbTextCompare) Then
        
            ' Just change the path after "DATABASE=" to be the path
            ' of the currently open database
            knownConnectionType = True
            databaseNameStart = InStr(ConnectString, "DATABASE=")
            ConnectStringFirstPart = Mid$(ConnectString, databaseNameStart)
            Debug.Print "Text  Link Old: " & ConnectString
            ConnectString = Left(ConnectString, databaseNameStart + Len("DATABASE=") - 1) & CurrentProject.Path
        
        ElseIf InStr(1, ConnectString, "Excel", vbTextCompare) Then

            ' Change the path after "DATABASE=" to be the path
            ' of the currently open database with the excel filename
            ' appended
            knownConnectionType = True
            databaseNameStart = InStr(ConnectString, "DATABASE=")
            ConnectStringFirstPart = Mid$(ConnectString, databaseNameStart)
            Debug.Print "Excel Link Old: " & ConnectString
            fileName = Mid(ConnectString, InStrRev(ConnectString, "\") + 1)
            ConnectString = Left(ConnectString, databaseNameStart + Len("DATABASE=") - 1) & CurrentProject.Path & "\" & fileName
        
        ElseIf InStr(1, ConnectString, ".accdb", vbTextCompare) Then
            
            ' Change the path after "DATABASE=" to be the path
            ' of the currently open database with the accdb filename
            ' appended
            knownConnectionType = True
            databaseNameStart = InStr(ConnectString, "DATABASE=")
            ConnectStringFirstPart = Mid$(ConnectString, databaseNameStart)
            Debug.Print "accdb Link Old: " & ConnectString
            fileName = Mid(ConnectString, InStrRev(ConnectString, "\") + 1)
            ConnectString = Left(ConnectString, databaseNameStart + Len("DATABASE=") - 1) & CurrentProject.Path & "\" & fileName
        
        End If
        
        ' If we detected one of the types above, set the link and
        ' refresh it
        If knownConnectionType Then
            tdf.Connect = ConnectString
            Debug.Print "      Link New: " & ConnectString
            tdf.RefreshLink
            Debug.Print ""
        End If
        
    Next tdf
End Function

Public Function UpdateCSVFolderLinks(Path As String)

' Example call: UpdateCSVFolderLinks("CSVdata")
    ' update connections to linked text files to the Path
    ' specified (relative to the current DB)

    Dim tdf As dao.TableDef
    Dim ConnectString As String
    Dim ConnectStringFirstPart  As String
    Dim databaseNameStart As Integer
    Dim knownConnectionType As Boolean
    Dim fileName As String

    ' Loop through all the items in the TableDefs collection
    ' and update for specific types
    For Each tdf In CurrentDb.TableDefs

        knownConnectionType = False
        ConnectString = tdf.Connect
        
        ' This routine only deals with Text file links
        If InStr(1, ConnectString, "Text", vbTextCompare) Then
        
            ' Just change the path after "DATABASE=" to be the path
            ' of the currently open database and the Path passed in
            knownConnectionType = True
            databaseNameStart = InStr(ConnectString, "DATABASE=")
            ConnectStringFirstPart = Mid$(ConnectString, databaseNameStart)
            Debug.Print "Text  Link Old: " & ConnectString
            ConnectString = Left(ConnectString, databaseNameStart + Len("DATABASE=") - 1) & _
                            CurrentProject.Path & "\" & _
                            Path
        End If
                            
        
        ' If we detected one of the types above, set the link and
        ' refresh it
        If knownConnectionType Then
            tdf.Connect = ConnectString
            Debug.Print "      Link New: " & ConnectString
            tdf.RefreshLink
            Debug.Print ""
        End If
        
    Next tdf
End Function


