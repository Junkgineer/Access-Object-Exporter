# Access-Object-Exporter
Exports Access Database forms, queries, and VBA code into more usable formats.

# Installation

This code itself is a VBA module. 
1. Open the Access database you wish to export.
2. Open the MS Access in-built Visual Basic editor
3. File->Import this module or create a new Module and paste in the code.

# Basic Usage

Set the `dbName` variable to a name you'd like it to be exported under. Usually this would be the database name or similar. This is primarily just for creating the directories and file names during the export. It does not need to match the actual database name verbatim. Ex:
`Public Const dbName = "MySuperAwesomeDatabase"`

Set the `export_dir` variable to the top level location for the objects to be exported to. It's recommended that this be an empty or nonexistant directory in order to avoid issues. If the directory does not exist, it'll be created. Ex:
`Public Const export_dir = "C:\Some\Export\Path\" & dbName & "\"`

To begin a full export, type "Fire_Everything" in the Immediate window and hit enter. You'll see each form opening, then closing, etc. Basic progress updates will be output to the Immediate window.

The files will be exported to the export directory within the folders "FORMS", "QRYS", and "VBA".

# Functions

**_Export_VBA()_** - Exports .cls, .frm, and .bas objects. These encompass the "Microsoft Access Class Objects", "Modules" and "Class Modules".
    
**_Export_Objects()_** - This exports all properties from each form. If it's made or set using "Design Mode" on a form, this will export it to a malformed JSON using the following structure:

    Base Form
        Properties  
        Objects
            Labels, Textboxes, etc
                Object Properties
        Sections
            acDetail          
                Section Properties          
                Objects
                    Labels, Textboxes, etc
                        Object Properties
            acFooter
                Section Properties
                Objects
                    Labels, Textboxes, etc
                        Object Properties
            acHeader
                Section Properties
                Objects
                    Labels, Textboxes, etc
                        Object Properties
            acPageFooter
                Section Properties
                Objects
                    Labels, Textboxes, etc
                        Object Properties
            acPageHeader
                Section Properties
                Objects
                    Labels, Textboxes, etc
                        Object Properties
        
An object property is exported whether it contains a value or not.

Some properties do not appear as they do in the form and need to be translated before effective use. Examples include colors, and object location coords (which is in Twips...because Microsoft)

The ControlType property defines what type of control it is (textbox, combo box, label, etc). The full list and translation of the value can be found in the Microsoft Access documentation here: https://docs.microsoft.com/en-us/office/vba/api/access.accontroltype

**_Export_Queries()_** - Exports the proper Queries saved in the database, and saves them as .sql files. It should be noted as an FYI that Access queries are not automatically compatible with any other SQL query, so simply running these queries in MySQL or MS-SQL will likely be met with errors.

# Extras

**_Validate_Data()_** - As an added bonus, this function checks the data in the database for a couple of common errors that are throwm when attempting to import Access databases into a MS-SQL server via the MS SQL Server Import and Export Data tool. 
    
**_CheckDates()_** - Access will let you enter an incorrectly formatted date into a date field unless told not to, wreaking havoc when a conventional SQL server is trying to read it. This won't FIX the issue, but it will tell you what table and column contains an error.

**_CheckLongvarCharCount()_** - There's no easy, automatic translation of super long text fields that Access allows into MS-SQL. This checks for fields containing more than 8000 characters of text...just in case someone had pasted in the contents of an entire email chain, for instance.

**_ListLinkedTables()_** - Outputs an easily digestible list in the immediate window of all the externally linked tables.

**_ConvertToLocal()_** - Converts the database to an entirely local version. Get's finniky with linked spreadsheets, but does pretty well all around.Iterates the linked tables, and sends each one to **_MakeLocalTable(tableName As String, Optional deleteOriginal As Boolean = True)_**

# Disclaimers

_I've used this on Access 2010 and 2016 without issue. It does require certain references be used, including Microsoft Visual Basic For Application Extensibility. If you're receiving errors that aren't explicitly obvious, there may be an issue with a reference._

_This was written to accommodate exporting multiple HUGE Access databases that would otherwise be far too time consuming. If you don't have many forms, queries, or modules in your database, it's probably faster and easier to just do it manually._

_Access databases are typically created by people who don't really know how to do proper SQL databases, and are ad-hoc. Microsoft recognizes this and makes Access very forgiving. Therefore, they can have all kinds of messy problems, including special characters in table/column names that would normally be illegal. There are a couple places in this module that attempts to fix these, but is by no means all-inclusive. Feel free to add checks of your own._

_This code is messy. I wrote it to use with a much larger application, and post processing occurs there. You will probably need to do the same, particularly with the VBA json's._

_There's a thousand ways to do this better, in a better language, but I needed it this way. Maybe you do too._
