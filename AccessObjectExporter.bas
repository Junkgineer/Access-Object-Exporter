Option Compare Database
' NOTE: None of these functions clean out their respective output dir's before writing to it.
' REQUIRED REFERENCES (Tools -> References, in the MS Access VB Editor):
'	Export_VBA() Requires: 				Microsoft Visual Basic For Application Extensibility' 
'	Export_Objects_Excel() Requires: 	Microsoft Excel 16.0 Object Library
' Drop this directly into a new module, and run Fire_Everything() from the immediate window or from a form.

' Name of the database (with no extension)
Public Const dbName = "MyDatabase"
'Directory where the exported files will be placed.
Public Const export_dir = "C:\Foo\Export\" & dbName & "\"
 
Public Sub Fire_Everything()
'Checks the primary parent export dir, then runs the standard export routines
    If (Len(Dir$(export_dir, vbDirectory)) > 0&) = False Then
        MkDir (export_dir)
    End If
    Debug.Print "Exporting to " & export_dir
    Export_Objects
    Export_Queries
    Export_VBA
    Debug.Print "All exports completed."
End Sub

Public Sub Export_VBA()
    Dim c As VBComponent
    Dim sfx As String
    Dim export_path As String
    Dim file_name As String
    Dim ctr As Integer
    Dim excise() As Variant
    
    excise = Array(" ", "~", "(", ")", """", "'", "/", ",", "?", "-")
    ctr = 0
    export_path = export_dir & "VBA\"
    If (Len(Dir$(export_path, vbDirectory)) > 0&) = False Then
        MkDir (export_path)
    End If
    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                sfx = ".cls"
            Case vbext_ct_MSForm
                sfx = ".frm"
            Case vbext_ct_StdModule
                sfx = ".bas"
            Case Else
                sfx = ""
            End Select
        file_name = c.Name
        If sfx <> "" Then
            For Each elem In excise
                If Mid(file_name, 1) = elem Then
                    file_name = Mid(file_name, 2)
                End If
                file_name = Replace(file_name, elem, "_")
                file_name = Replace(file_name, "$", "__DOLLAR__")
            Next
                c.Export _
                    FileName:=export_path & file_name & sfx
        End If
        ctr = ctr + 1
    Next
    Debug.Print ctr & " VBA modules and classes exported."
End Sub
 
Public Sub Export_Objects()
' Exports all the objects and their properties on every form.
    Dim frm As Form
    Dim FormName As String
    Dim sectionName As String
    Dim frm_ctr As Integer
    Dim ctrl As Control
    Dim ctrl_ctr As Integer
    Dim frm_prop_ctr As Integer
    Dim prop_ctr As Integer
    Dim sect_ctr As Integer
    Dim obj_ctr As Integer
    Dim form_prop As Property
    Dim form_value As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Dim export_path As String
    export_path = export_dir & "FORMS\"
    If (Len(Dir$(export_path, vbDirectory)) > 0&) = False Then
        MkDir (export_path)
    End If
    
    For frm_ctr = 0 To CurrentProject.AllForms.Count - 1
        If Not CurrentProject.AllForms(frm_ctr).IsLoaded Then
            DoCmd.OpenForm CurrentProject.AllForms(frm_ctr).Name, acDesign
        End If
        Set frm = Forms(CurrentProject.AllForms(frm_ctr).Name)
        FormName = CurrentProject.AllForms(frm_ctr).Name
        Set oFile = fso.CreateTextFile(export_path & FormName & ".json")
        
        oFile.WriteLine ("{ """ & FormName & """: [{")
        frm_prop_ctr = 1
        For Each form_prop In frm.Properties
            On Error Resume Next
            form_value = Replace(form_prop.Value, " ", "")
            form_value = Replace(form_value, """", "")
            oFile.WriteLine ("""" & form_prop.Name & """: """ & form_value & """,")
            frm_prop_ctr = frm_prop_ctr + 1
        Next
        
        oFile.WriteLine ("""Sections"": [{")
        
            sect_ctr = 0
           oFile.WriteLine ("""acDetail"" : [{")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.Count
                form_value = Replace(frm.Section(sect_ctr).Properties(prop_ctr), " ", "")
                form_value = Replace(form_value, """", "")
                If prop_ctr = frm.Section(sect_ctr).Properties.Count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """,")
                End If
            Next
            oFile.WriteLine ("}],")
            
            sect_ctr = 1
            oFile.WriteLine ("""acHeader"" : [{")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.Count
                form_value = Replace(frm.Section(sect_ctr).Properties(prop_ctr), " ", "")
                form_value = Replace(form_value, """", "")
                If prop_ctr = frm.Section(sect_ctr).Properties.Count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """,")
                End If
            Next
            oFile.WriteLine ("}],")
            
            sect_ctr = 2
            oFile.WriteLine ("""acFooter"" : [{")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.Count
                form_value = Replace(frm.Section(sect_ctr).Properties(prop_ctr), " ", "")
                form_value = Replace(form_value, """", "")
                If prop_ctr = frm.Section(sect_ctr).Properties.Count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """,")
                End If
            Next
            oFile.WriteLine ("}],")
            
            sect_ctr = 3
            oFile.WriteLine ("""acPageHeader"" : [{")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.Count
                form_value = Replace(frm.Section(sect_ctr).Properties(prop_ctr), " ", "")
                form_value = Replace(form_value, """", "")
                If prop_ctr = frm.Section(sect_ctr).Properties.Count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """,")
                End If
            Next
            oFile.WriteLine ("}],")
            
            sect_ctr = 4
            oFile.WriteLine ("""acPageFooter"" : [{")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.Count
                form_value = Replace(frm.Section(sect_ctr).Properties(prop_ctr), " ", "")
                If prop_ctr = frm.Section(sect_ctr).Properties.Count - 1 Then
                    form_value = Replace(frm.Section(sect_ctr).Properties(prop_ctr), " ", "")
                    form_value = Replace(form_value, """", "")
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & form_value & """,")
                End If
            Next
            oFile.WriteLine ("}]")
        oFile.WriteLine ("}],")
        oFile.WriteLine ("""Objects"": [{")
        obj_ctr = 0
        For Each ctrl In frm.Controls
            On Error Resume Next
                oFile.WriteLine ("""" & ctrl.Name & """: [{")
                For prop_ctr = 0 To ctrl.Properties.Count
                    form_value = Replace(ctrl.Properties(prop_ctr), " ", "")
                    form_value = Replace(form_value, """", "")
                    If prop_ctr = ctrl.Properties.Count - 1 Then
                        oFile.WriteLine ("""" & ctrl.Properties(prop_ctr).Name & """: """ & form_value & """")
                    Else
                        oFile.WriteLine ("""" & ctrl.Properties(prop_ctr).Name & """: """ & form_value & """,")
                   End If
                Next
                If obj_ctr = frm.Controls.Count - 1 Then
                    oFile.WriteLine ("}]")
                Else
                    oFile.WriteLine ("}],")
                End If
                obj_ctr = obj_ctr + 1
                Debug.Print (obj_ctr + " " + frm.Controls.Count)
        Next
        oFile.WriteLine ("}]")
        oFile.WriteLine ("}]")
        oFile.WriteLine ("}")
        If frm.Name <> "frmReporter" Then
            DoCmd.Close acForm, frm.Name
        End If
    Next
        
    Set fso = Nothing
    Set oFile = Nothing
    Debug.Print frm_ctr & " forms exported."
 
End Sub
 
Public Sub Export_Queries()
' Exports all the queries to SQL text files.
' Cleans most of the garbage out of the query names (who puts slashes in file names????)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Dim db As DAO.Database
    Dim qrydef As DAO.QueryDef
    Dim qryName As String
    Dim excise() As Variant
    Dim elem As Variant
    Dim ctr As Integer
    Dim export_path As String
    export_path = export_dir & "QRYS\"
    'Check if the export directory exists, create it if False:
    If (Len(Dir$(export_path, vbDirectory)) > 0&) = False Then
        MkDir (export_path)
    End If
    
    excise = Array(" ", "(", ")", """", "'", "/", ",", "?", "-")
    ctr = 1
    
    Set db = CurrentDb()
    
    For Each qrydef In db.QueryDefs
        qryName = qrydef.Name
        On Error GoTo err
        For Each elem In excise
            qryName = Replace(qryName, elem, "_")
            qryName = Replace(qryName, "$", "__DOLLAR__")
        Next
        'Debug.Print "Exporting Query " & qryName
        Set oFile = fso.CreateTextFile(export_path & qryName & ".sql")
        oFile.Write qrydef.SQL
        ctr = ctr + 1
    Next qrydef
 
    Set qrydef = Nothing
    Set db = Nothing
    Set fso = Nothing
    Set oFile = Nothing
    Debug.Print ctr & " queries exported."
    
err:
    Debug.Print "Error exporting query: " & qryName
    Resume Next
    
End Sub
 
Public Sub ValidateData()
'This does NOT get run as part of the normal export. It's used to ready the data
' to be used with the SQL Server Import/Export tool.
 
' Runs queries on each table looking for common issues that result in errors
' when trying to export the data to SQL via the SQL Server Data Import/Export tool.
' Improper dates in a field are the most common, followed by Longtext fields exceeding
' the maximum character count for VARCHAR(MAX), which is 8000. Additional tests
' can be added as necessary. This function iterates everything, then sends
' the columns off to perform the apropriate test in individual subs.
 
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
 
    Set db = CurrentDb
    Debug.Print "Checking..."
    For Each tdf In db.TableDefs
        If Not (tdf.Name Like "Msys*" Or tdf.Name Like "~*") Then
            For Each fld In tdf.Fields
                If fld.Type = 8 Then 'Date fields
                    CheckDates tdf.Name, fld.Name
                ElseIf fld.Type = 12 Then
                    CheckLongvarCharCount tdf.Name, fld.Name 'LongText fields
                End If
            Next
        End If
    Next
    Set tdf = Nothing
    Set db = Nothing
    Debug.Print "Done"
End Sub
 
Public Sub CheckDates(TblName As String, fldName As String)
' Checks for malformed dates that MS-SQL would not allow.
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    
    Set dbs = CurrentDb
    dateQry = "SELECT * FROM [" & TblName & "] WHERE (IsDate([" & fldName & "]) = False " & _
        "And Len([" + fldName + "]) > 0) " & _
        "Or [" & fldName & "] < 01/01/1990"
    Set rs = dbs.OpenRecordset(dateQry)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Debug.Print "[" & TblName & "].[" & fldName & "] invalid date: " & rs.Fields(fldName)
            rs.MoveNext
        Loop
        rs.Close
    End If
    
End Sub
 
Public Sub CheckLongvarCharCount(TblName As String, fldName As String)
' Checks Longvar column types for strings exceeding 8000 chars, which is the maximum MS-SQL
'   allows (in the type the auto converter uses for Longvar)
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    
    Set dbs = CurrentDb
    charCntQry = "SELECT * FROM [" & TblName & "] WHERE Len([" & fldName & "]) > 8000"
    Set rs = dbs.OpenRecordset(charCntQry)
    If rs.RecordCount > 0 Then
        
        Do While Not rs.EOF
            Debug.Print "[" & TblName & "].[" & fldName & "] exceeds max length: " & rs.Fields(fldName)
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub
 
Public Sub Export_Objects_Excel()
' Exports a list of all forms and objects into a spreadsheet
'       for easier user consumption
' Each form is given it's own sheet.
 
    Dim frm As Form
    Dim FormName As String
    Dim frm_ctr As Integer
    Dim ctrl As Control
    Dim ctrl_ctr As Integer
    Dim prop_ctr As Integer
    Dim form_prop As Property
    Dim form_prop_ctr As Integer
    
    Dim appExcel As Excel.Application
    Set appExcel = CreateObject("excel.Application")
    
    Dim wbExcel As Excel.Workbook
    Set wbExcel = appExcel.Workbooks.Add
    Set wbExcel = appExcel.ActiveWorkbook
    appExcel.Visible = True
    Set wsExcel = wbExcel.ActiveSheet
    
    For frm_ctr = 0 To CurrentProject.AllForms.Count - 1
        
        If Not CurrentProject.AllForms(frm_ctr).IsLoaded Then
            DoCmd.OpenForm CurrentProject.AllForms(frm_ctr).Name, acDesign
        End If
        Set frm = Forms(CurrentProject.AllForms(frm_ctr).Name)
        FormName = CurrentProject.AllForms(frm_ctr).Name
        
        Debug.Print ("Processing " & frm_ctr & " of " & CurrentProject.AllForms.Count & ": " & FormName)
        
        Set wsExcel = wbExcel.Worksheets.Add
        form_prop_ctr = 1
        wsExcel.Name = Left(FormName, 31)
        wsExcel.Cells(form_prop_ctr, "A") = "FORM NAME"
        wsExcel.Cells(form_prop_ctr, "B") = FormName
        form_prop_ctr = form_prop_ctr + 1
        
        For Each form_prop In frm.Properties
            On Error Resume Next
            wsExcel.Cells(form_prop_ctr, "A") = form_prop.Name
            wsExcel.Cells(form_prop_ctr, "B") = form_prop.Value
            form_prop_ctr = form_prop_ctr + 1
        Next
        ctrl_ctr = 1
        For Each ctrl In frm.Controls
            On Error Resume Next
                wsExcel.Cells(form_prop_ctr, "A") = "CONTROL"
                wsExcel.Cells(form_prop_ctr, "B") = ctrl_ctr & " OF " & frm.Controls.Count
                form_prop_ctr = form_prop_ctr + 1
                ctrl_ctr = ctrl_ctr + 1
                For prop_ctr = 0 To ctrl.Properties.Count
                    wsExcel.Cells(form_prop_ctr, "A") = ctrl.Properties(prop_ctr).Name
                    wsExcel.Cells(form_prop_ctr, "B") = ctrl.Properties(prop_ctr)
                    form_prop_ctr = form_prop_ctr + 1
                Next
        Next
        
        If frm.Name <> "frmReporter" Then
            DoCmd.Close acForm, frm.Name
        End If
        
    Next
 
End Sub
Sub ListLinkedTables()
'Just an extra tool for getting an easy list of all the externally linked tables.
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    
    Set dbs = CurrentDb
    linkedTableQry = "SELECT * FROM MSysObjects WHERE [ForeignName] IS NOT NULL;"
    Set rs = dbs.OpenRecordset(linkedTableQry)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Debug.Print "Local Name: " & rs.Fields("Name") & _
            " | Foreign Name: " & rs.Fields("ForeignName") & _
            " | Path: " & rs.Fields("Database")
            rs.MoveNext
        Loop
        rs.Close
    End If
    Debug.Print "Done"
 
End Sub
Sub ListLinkedTablesCSV()
'Same as above, but exports to a CSV.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim csvList As String
    
    Set dbs = CurrentDb
    linkedTableQry = "SELECT * FROM MSysObjects WHERE [ForeignName] IS NOT NULL;"
    Set rs = dbs.OpenRecordset(linkedTableQry)
    If rs.RecordCount > 0 Then
        Set oFile = fso.CreateTextFile(export_dir & dbName & "-linked_tables.csv")
        oFile.WriteLine ("Local Name, Foreign Name, Path")
        Do While Not rs.EOF
            oFile.WriteLine (rs.Fields("Name") & "," & _
                rs.Fields("ForeignName") & "," & _
                rs.Fields("Database"))
                rs.MoveNext
        Loop
        rs.Close
    End If
    Debug.Print "Wrote file to: " & export_dir & dbName & "-linked_tables.csv"
    
    Set fso = Nothing
    Set oFile = Nothing
    
    Debug.Print "Done"
 
End Sub
Sub ConvertToLocal()
' Another extra tool. Converts the database to a single, unlinked database if it had any external tables.
' There are a few things it doesn't like, such as linked spreadsheets, but will handle most everything else.
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    
    Set dbs = CurrentDb
    linkedTableQry = "SELECT * FROM MSysObjects WHERE [ForeignName] IS NOT NULL;"
    Set rs = dbs.OpenRecordset(linkedTableQry)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            If rs.Fields("Name") <> "TBL Contract Data" Then
                Debug.Print "Converting to Local Table: " & rs.Fields("Name")
                MakeTableLocal (rs.Fields("Name"))
            Else
                Debug.Print "Skipping Table: " & rs.Fields("Name")
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Debug.Print "Conversion Complete"
 
End Sub
 
Sub MakeTableLocal(tableName As String, Optional deleteOriginal As Boolean = True)
' ConvertToLocal iterates the tables, this converts them.
    Dim DbPath As Variant, TblName As Variant
 
    DbPath = DLookup("Database", "MSysObjects", "Name='" & tableName & "' And Type=6")
    TblName = DLookup("ForeignName", "MSysObjects", "Name='" & tableName & "' And Type=6")
    If IsNull(DbPath) Then
        Exit Sub
    End If
 
    If deleteOriginal Then
        DoCmd.DeleteObject acTable, tableName
    Else
        tableName = tableName & " - local"
    End If
    
    DoCmd.TransferDatabase acImport, "Microsoft Access", DbPath, acTable, TblName, tableName
End Sub
