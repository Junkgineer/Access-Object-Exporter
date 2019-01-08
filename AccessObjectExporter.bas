Option Compare Database
' REQUIRES REF: Microsoft Visual Basic For Application Extensibility'

Public Const dbName = "MSD"
Public Const export_dir = "C:\Some\Export\Path\" & dbName & "\"

Public Sub Fire_Everything()
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
    Dim ctr As Integer
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

        If sfx <> "" Then
            c.Export _
                Filename:=export_path & c.Name & sfx
        End If
        ctr = ctr + 1
    Next

    Debug.Print ctr & " VBA objects exported."
End Sub

 

Public Sub Export_Objects()
' Exports all the objects and their properties on every form.
' Creates a malformed JSON, but it's close.
    Dim frm As Form
    Dim FormName As String
    Dim sectionName As String
    Dim frm_ctr As Integer
    Dim ctrl As Control
    Dim ctrl_ctr As Integer
    Dim frm_prop_ctr As Integer
    Dim prop_ctr As Integer
    Dim sect_ctr As Integer
    Dim form_prop As Property

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Dim export_path As String
    export_path = export_dir & "FORMS\"
    'Check if the export directory exists, create it if False:

    If (Len(Dir$(export_path, vbDirectory)) > 0&) = False Then
        MkDir (export_path)
    End If

    For frm_ctr = 0 To CurrentProject.AllForms.count - 1
        If Not CurrentProject.AllForms(frm_ctr).IsLoaded Then
            DoCmd.OpenForm CurrentProject.AllForms(frm_ctr).Name, acDesign
        End If

        Set frm = Forms(CurrentProject.AllForms(frm_ctr).Name)
        FormName = CurrentProject.AllForms(frm_ctr).Name
        Set oFile = fso.CreateTextFile(export_path & FormName & ".json")

        oFile.WriteLine ("{ """ & FormName & """: {")
        frm_prop_ctr = 1
        For Each form_prop In frm.Properties
            On Error Resume Next
            If frm_prop_ctr = frm.Properties.count Then
                oFile.WriteLine ("""" & form_prop.Name & """: """ & form_prop.Value & """")
            Else
                oFile.WriteLine ("""" & form_prop.Name & """: """ & form_prop.Value & """,")
            End If
            frm_prop_ctr = frm_prop_ctr + 1

        Next

        oFile.WriteLine ("""Sections"": {")

            sect_ctr = 0
            oFile.WriteLine ("""acDetail"" : {")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.count
                If prop_ctr = frm.Section(sect_ctr).Properties.count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """,")
                End If
            Next
            oFile.WriteLine ("},")
            sect_ctr = 1
            oFile.WriteLine ("""acHeader"" : {")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.count
                If prop_ctr = frm.Section(sect_ctr).Properties.count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """,")
                End If
            Next
            oFile.WriteLine ("},")
            sect_ctr = 2
            oFile.WriteLine ("""acFooter"" : {")

            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.count
                If prop_ctr = frm.Section(sect_ctr).Properties.count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """,")
                End If
            Next
            oFile.WriteLine ("},")

            sect_ctr = 3
            oFile.WriteLine ("""acPageHeader"" : {")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.count
                If prop_ctr = frm.Section(sect_ctr).Properties.count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """,")
                End If
            Next

            oFile.WriteLine ("},")
            sect_ctr = 4
            oFile.WriteLine ("""acPageFooter"" : {")
            For prop_ctr = 0 To frm.Section(sect_ctr).Properties.count
                If prop_ctr = frm.Section(sect_ctr).Properties.count - 1 Then
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """")
                Else
                    oFile.WriteLine ("""" & frm.Section(sect_ctr).Properties(prop_ctr).Name & """: """ & frm.Section(sect_ctr).Properties(prop_ctr) & """,")
                End If
            Next
           oFile.WriteLine ("},")
        oFile.WriteLine ("},")
        oFile.WriteLine ("""Objects"": {")

        For Each ctrl In frm.Controls
            On Error Resume Next
                oFile.WriteLine ("""" & ctrl.Name & """: {")
                For prop_ctr = 0 To ctrl.Properties.count
                    If prop_ctr = ctrl.Properties.count - 1 Then
                        oFile.WriteLine ("""" & ctrl.Properties(prop_ctr).Name & """: """ & ctrl.Properties(prop_ctr) & """")
                    Else
                        oFile.WriteLine ("""" & ctrl.Properties(prop_ctr).Name & """: """ & ctrl.Properties(prop_ctr) & """,")
                    End If
                Next
                oFile.WriteLine ("},")
        Next
        oFile.WriteLine ("}")
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
' Cleans most of the garbage out of the query names (who the @#%$@ puts slashes in file names????)

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

Public Sub CheckDates(tblName As String, fldName As String)
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Set dbs = CurrentDb

    dateQry = "SELECT * FROM [" & tblName & "] WHERE (IsDate([" & fldName & "]) = False " & _
        "And Len([" + fldName + "]) > 0) " & _
        "Or [" & fldName & "] < 01/01/1990"

    Set rs = dbs.OpenRecordset(dateQry)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Debug.Print "[" & tblName & "].[" & fldName & "] invalid date: " & rs.Fields(fldName)
            rs.MoveNext
        Loop
        rs.Close
    End If

End Sub

Public Sub CheckLongvarCharCount(tblName As String, fldName As String)
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset

    Set dbs = CurrentDb
    charCntQry = "SELECT * FROM [" & tblName & "] WHERE Len([" & fldName & "]) > 8000"
    Set rs = dbs.OpenRecordset(charCntQry)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Debug.Print "[" & tblName & "].[" & fldName & "] exceeds max length: " & rs.Fields(fldName)
            rs.MoveNext
        Loop
        rs.Close
    End If
End SubS
