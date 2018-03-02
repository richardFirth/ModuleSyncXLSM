Attribute VB_Name = "AAA_Global"
Option Explicit


' allows the macro to track if the ModuleUpdater UI is open or closed.
' we don't want multiple Presses of the "Run Module Sync tool" button to open multiple UI's
Public UI_OPEN As Boolean

Public Const StoreFilesOnDesktop As Boolean = False ' to store the module file structure on the desktop

Public Const BestModules As String = "BestModules.xlsm"

' Public Type for ModuleVersionData.
' This represents the header data for a single module, and is used to compare different modules

 '$-VERSIONCONTROL
 '$-*MINOR_VERSION*1.0
 '$-*DATE*18Jan18
 '$-*NAME*example



Public Type ModuleVersionData
    A_Name As String
    B_ModulePath As String
    C_Vcontrol As Boolean
    
    D_ID As String
    E_MajorVersion As String
    F_MinorVersion As String
    G_date As String

    H_OldVersion As Boolean
    I_newModule As Boolean
    J_CodeChange As Boolean
End Type

Public Enum ModuleSpecies
    A_NORMAL
    B_CLASS
    C_FORM
End Enum


Sub EraseExportedFolder()
    'erases the folder in the temp area
    Dim tFolder As String: tFolder = folderToPlaceData & "\ModuleSyncOutput"
    If FolderThere(tFolder) Then Call DeleteFolderTreeRF(tFolder)
    
End Sub

Function folderToPlaceData() As String
    ' fetches the folder path to deposit module data
    If StoreFilesOnDesktop Then
        folderToPlaceData = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    Else
        folderToPlaceData = Environ("TEMP")
    End If
End Function


Sub testdate()

Dim dTest As String: dTest = "11Feb2018"

 '   Dim bestDate As Date: bestDate = DateValue(dTest)

    MsgBox Date

End Sub
