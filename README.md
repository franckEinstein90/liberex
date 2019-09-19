Librerex – A content inventorying system for Parks Canada

User Documentation
Storing File Related information
File Information Storage in the Database

 
Folder Information - tblFolders
Folder-specific information is stored in the tblFolders table:
Fields	Keying		
ID	Primary Key		
FolderName			
ParentFolderID	Foreign Key (tblFolders)		The ID of the parent folder for this folder
SharedDriveID	Foreign Key (tlbSharedDrives)		
BusinessOwnerID	Foreign Key (tblBusinessOwners)		

File Information - tblFiles
File specific information is stored in the tblFiles table: 
Field Name		Type
ID	Primary Key	
FolderID	Foreign Key (tblFolders)	
FileName		String
FileType	Foreign Key (…)	
Author		
OwnerID		
CreationDate		Date/Time
LastAccess		Date/Time
LastChanged		Date/Time
Size		Number
ImportID	Foreign Key	Num


Producing New Scans for Input
Scans are produced via the TreeSize Pro tool. Librerex accepts TSP output of type: 
-	txt
-	Excel
-	Csv
Each type of scan is associated with a scan format, specifying the names and positions of the columns included in the scan. For example, a text duplicate scan is produced in this format by TSP: 
 
With each column and each column value separated by a tab. In the database, the “Duplicates” scan type is associated with the following format items (stored in the tblScanFormats table):
 
Therefore, the system is able to “understand” the scan output and handle the migration into its database. 
When importing data, Librerex will look for a header line with column labels in the first 10 rows. If it is unable to find such a matching header, it will abort the migration. 
Inserting Information contained in an Existing Data Source 
Process of grabbing data

 

Data Grab From
Data grabs are initiated by the user via the fldlgAbsorbContent form
 
Each new data grab is associated with an ID, which is the id of a new record in the tblInputSheets table. Each new piece of information added to the data base via the grab is associated with that grab ID
 
In the database, the relationship between files, folders, and data grab is modeled as: 
 

Data Sources
A data source is a file that contains an information asset inventory. This system accepts the following formats: 
Data Source Formats:
-	Text (tab separated)
-	csv 
-	excel
A data source file must include a header with row labels indicating the columns included. Librerex will start reading input after finding the header, and matching it against the profile of the data source as specified by the user. 
 
See the Producing New Scans section for instructions on how to create new scans ready for integration. 
1.	From the sys admin form, navigate to “Insert Existing Dataset”
 
2.	The flgAbsorbContent form is used to insert new tsp scans into the database – the user can use the form to select a data source, associate it with a given Server/Shared Drive combination and designate a business owner. 
 

The system accepts scans in the following format: 
-	Text
-	Excel
-	CSV
Each scan type has a given profile, and the user can select the appropriate profile using the choice item included for that purpose. 
Each new import will create an entry in the … table, indicating: 
-	The date and time of the start of the import
-	The number of files and folders concerned by the import
From there, the user can select the type of scan to be imported into the main database, as well as other relevant information, such as the location of the scan, the server concerned by the scan, the Shared Drive concerned by the scan, as well as the business owner of the scan.

 


Scan Profiles
A scan profile specifies the columns and format associated with either a given file or a group of file. Scan profiles are stored in the tblTSPScanTypes table and in the tblScanFormats table.


 
Data source types - tblScanTypes 

 
GroupedData
A record of type TSPScanType can be grouped. This applies to scans that group files according to some criteria. For instance, in the TSP scan below, exact versions of the same file are grouped under a header. 
 
When that data is read by the system, it needs to be aware of the 
Currently, there are only two predefined scans set, one for drafts, and one for duplicates. 
Fields		
		
GroupedData		Indicates if this is a grouped scan or not


 
of scans

 
Business Owner Information – tblBusinessOwners
The primary purpose of the system is to produce reports to help business owners make decisions in relation to their information assets. Business owner information is stored in the tblBusinessOwners table, and business owners are associated with information assets at the folder level through the grabs that occurred and that concerned them. 
 


 
Developer Documentation
System Architecture





Reading File Information
When reading from a data source, the system stores each read record into a translating structure called structInputFileInfo, which is used by the internals of the system to create or update records in tblFiles and in tblFolders. The debug screen shot below displays the values associated with the structure
 
The processFileRecord method of the objSharedDrive object takes an object of type structInputFileInfo as argument, and uses the information contained in it to populate or update the database. 

Private Sub processFileRecord(ByRef values As structInputFileInfo, _
ByRef dataSource As objDataSource)
Dim folder_id As Long
folder_id = process_folder_path(values.fields("Path"), dataSource.m_grab_id)
…
Appendices
VBA Modules
01 – Low Level Routines
General low level utilities used throughout the system
Public Function isFolderPath(ByVal path As String) As Boolean	
Returns true is the path passed as argument corresponds to the path of a path. For example: 
C:\Folder1\Folder2\
Public Function strDBProcess(ByVal word As String) As String
Processes a string by quoting characters that need to be escaped for SQL queries. For example, ‘ becomes ‘’ 
, as long


Public Function stringFileSizeToNumber(ByVal inp As String) As Double
Converts file size information provided as a string such as “65.5 kb” into a numerical value (for example, 65.5)

02 – Low Level Database Routines

Public Function getFolderID(		
	ByVal path as String		
	ByVal ShareID as Long		
	) As Long		
Looks for a folder ID in the database, based on the path supplied as argument, returns -1 if the folder associated with the path doesn’t exist

Public Function countSubfolders(ByVal folderID As Long) As Long
Counts the number of subfolders of the folder who's id is passed as argument

Public Function countFiles(ByVal folderID As Long) As Long
Counts the number of files contained in the folder who's id is passed as argument

Public Function registerFolderPath(	ByVal ShareID As Long	
	ByVal path as string) As Long
Adds folder record to db for all the folders in path arguments when they don't exist. Returns ID of the folder corresponding to the whole path

03 – Debugging and Error Handling
Contains routine and global variables for debugging and error handling
Global Variables
Global glb_debug_on As Boolean	When on, extra debug info is displayed, either in the immediate window, or through the user interface. As well, some breakpoints get set

VBA Objects
objComputerFile
Encapsulates information related to a file - used for data sources. Can encapsulate information related to a file stored on:
-	The local file server
-	A SharePoint site (to do)
Data
m_fileExists as Boolean	
Methods
Public Sub initByPath(ByVal full_path As String)
Initializes a new object of type objComputerFile based on the path passed as argument. Checks if the file exists, and assigns the variable m_fileExists correspondingly

objDatabaseRecord
Methods
Initialization
Public Sub initByID(ByVal table_name As String, ByVal obj_id As Long)
The object must already exist in the database for this method to be applicable

Public Function getRecord() as dao.recordset
Returns a record set corresponding to the object – object must already exist in the database

objDataSource
Encapsulates information related to a data source. 
Data
m_dbRecord As objDatabaseRecord	
m_sourceFile As objComputerFile	
Public m_colsFormat As Scripting.dictionary	Contains the indexes of each columns in the data source, and their label
Public m_dataGrabID As Long	Each data source is associated with a grab id when consumed by the system. When a file or a folder the system is read or updated based on a data source, that file or folder is associated with the grab_id of the read. If information related to a file or folder is updated, then the grab_id remains 
Public m_businessOwnerID As Long	The ID of the business owner associated with the grab
		

Methods
Initialization
Public Sub initByName(ByVal scanTypeName As String, ByVal scanFilePath As String)
Initializes a new object of type objDataSource. The first argument is the name of the type of scan of the data source (for example “Draft” or “Duplicates”) and the second argument is the path of the data source – All possible scan types are named and defined in the tblScanTypes table


Private Sub makeDataSourceSignature()
Builds the signature of the data source based on the information contained in the database, stores the template in the m_colsFormat variable

Public Sub prepareNewDataGrab(ByRef share As objSharedDrive)
Prepares a new data grab – Creates a new entry in the tblInputSheets table, and populates the m_dataGrabID with the ID of the new entry 

objFileRecord
Encapsulates information related to a given file in the repository
Data
Public m_folderID as Long	The ID of the folder record in which the file is stored

Methods
Initialization
Public Sub initByName(ByVal file_name As String, ByVal folder_id As Long)
Initializes a new file record based on the name and the folder_id passed as argument

objFolderRecord
Encapsulate information related to a givne folder
Data
m_dataGrabID As Long	Each data source is associated with a grab id when consumed by the system. When a file or a folder the system is read or updated based on a data source, that file or folder is associated with the grab_id of the read. If information related to a file or folder is updated, then the grab_id remains 
		
		

Methods
Private Function createNewFolderRecordFromPath (		
	ByRef SD as ojbSharedDrive,		
	ByVal Path As String)		
	As Long		
Adds folder record to db for all folders in path arguments when they don't already exist. Returns ID of the folder corresponding to path

Private Sub initByPath(		
	ByRef SD as ojbSharedDrive,		
	ByVal Path As String)		
Initializes a new instance based on the path supplied as argument 

objSharedDrive
Encapsulates information related to a particular share drive. 
Data
m_dbRecord As objDatabaseRecord	

Methods
Public sub initByID(ByVal s_id as Long)		
Initializes a new instance based on the ID supplied as argument 

Public Function getID() As Long		
Returns the identifier of the share drive 

Private Function processFolderPath(		
	ByRef values As structInputFileInfo,		
	ByRef dataSource As objDataSource)		
	As Long		
Processes a folder path, adding the relevant folders to the database if they're not in already. Returns the id of the folder in the database.

Private Sub processFileRecord(		
	ByRef values As structInputFileInfo,		
	ByRef dataSource As objDataSource)		
			



Structures
structInputFileInfo 
(See Reading File Information for further explanations)

