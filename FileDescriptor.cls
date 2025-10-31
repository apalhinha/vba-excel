VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File descriptor for use with FolderScan

Public Domain As String         ' auto, the first folder after LocalRoot
Public Category As String       ' auto, the second folder after LocalRoot
Public Folder As String         ' auto, the third folder after LocalRoot
Public ObjectType As String     ' auto, eg. 'PowerPoint'
Public fileName As String       ' auto, as seen from Parent Folder. Example 'Conference_Template.pptx' or 'Subfolder name'
Public ObjectName As String     ' auto, pathname after Folder. Example 'Conference.pptx' or 'Subfolder name' or 'Subfolder/Inner/Somefile.pptx'
Public RelativePath As String   ' auto, path and filename staring from the LocalRoot folder. Eg., "Admin\2503 Human Resources\Trainning\Some file.xlsx"

' These fields are maintained in Excel and we will addressthem later:
Public DateFound As Date        ' The date the file was found
Public Behaviour As String      ' Maintained in Excel. ShowNothing, ShowFolders, ShowEverithing. Used for Folders only
Public Stereotype As String     ' Example: Policy, Project Plan
Public Reference As String      ' Example: AA/BB/CC
Public Link As String           ' Example: http://...
Public Description As String    ' Example: Text description
Public Keywords As String       ' Example: .Keyword1.Keyword2.Key word 3.


