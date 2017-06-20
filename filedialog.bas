Attribute VB_Name = "filedialog"
Option Explicit

'Select files – msoFileDialogFilePicker
'select file filedialogThe msoFileDialogFilePicker dialog type allows you to select one or more files.
'
'Select single files
'
'The most common select file scenario is asking the user to select a single file. The code below does just that:

Dim fDialog As filedialog, result As Integer
Set fDialog = Application.filedialog(msoFileDialogFilePicker)
     
'Optional: FileDialog properties
fDialog.AllowMultiSelect = False
fDialog.Title = "Select a file"
fDialog.InitialFileName = "C:\"
'Optional: Add filters
fDialog.Filters.Clear
fDialog.Filters.Add "Excel files", "*.xlsx"
fDialog.Filters.Add "All files", "*.*"
 
'Show the dialog. -1 means success!
If fDialog.Show = -1 Then
   Debug.Print fDialog.SelectedItems(1)
End If
'The result can look similar to this:
'
'
'C:\somefile.xlsx
'Select multiple files
'
'Quite common is a scenario when you are asking the user to select one or more files. The code below does just that. Notice that you need to set AllowMultiSelect to True.

Dim fDialog As filedialog, result As Integer
Set fDialog = Application.filedialog(msoFileDialogFilePicker)
     
'IMPORTANT!
fDialog.AllowMultiSelect = True
 
'Optional FileDialog properties
fDialog.Title = "Select a file"
fDialog.InitialFileName = "C:\"
'Optional: Add filters
fDialog.Filters.Clear
fDialog.Filters.Add "Excel files", "*.xlsx"
fDialog.Filters.Add "All files", "*.*"
 
'Show the dialog. -1 means success!
If fDialog.Show = -1 Then
  For Each it In fDialog.SelectedItems
    Debug.Print it
  Next it
End If
'The result can look similar to this:

'C:\somefile.xlsx
'C:\somefile1.xlsx
'C:\somefile2.xlsx
'Select folder – msoFileDialogFilePicker
'select folder application.filedialogSelecting a folder is more simple than selecting files. However only a single folder can be select within a single dialog window.

Set fDialog = Application.filedialog(msoFileDialogFolderPicker)
 
'Optional: FileDialog properties
fDialog.Title = "Select a folder"
fDialog.InitialFileName = "C:\"
 
If fDialog.Show = -1 Then
  Debug.Print fDialog.SelectedItems(1)
End If
'The msoFileDialogFolderPicker dialog allows you to only select a SINGLE folder and obviously does not support file filders
'Open file – msoFileDialogOpen
'file open application.filedialogOpening files is much more simple as it usually involves a single file. The only difference between the behavior between Selecting and Opening files are button labels.
'
'The open file dialog will in fact not open any files! It will just allow the user to select files to open. You need to open the files for reading / writing yourself. Check out my posts:
'Read file in VBA
'Write file in VBA
'Open file example
'
'The dialog below will ask the user to select a file to open:

Dim fDialog As filedialog, result As Integer, it As Variant
Set fDialog = Application.filedialog(msoFileDialogOpen)
     
'Optional: FileDialog properties
fDialog.Title = "Select a file"
fDialog.InitialFileName = "C:\"
     
'Optional: Add filters
fDialog.Filters.Clear
fDialog.Filters.Add "All files", "*.*"
fDialog.Filters.Add "Excel files", "*.xlsx"
   
If fDialog.Show = -1 Then
  Debug.Print fDialog.SelectedItems(1)
End If
'Save file – msoFileDialogSaveAs
'saveas application.filedialogSaving a file is similarly easy, and also only the buttons are differently named.
'
'The save file dialog will in fact not save any files! It will just allow the user to select a filename for the file. You need to open the files for reading / writing yourself. Check out my post on how to write files in VBA
'Save file example
'
'The dialog below will ask the user to select a path to which a files is to be saved:

Dim fDialog As filedialog, result As Integer, it As Variant
Set fDialog = Application.filedialog(msoFileDialogSaveAs)
 
'Optional: FileDialog properties
fDialog.Title = "Save a file"
fDialog.InitialFileName = "C:\"
 
If fDialog.Show = -1 Then
  Debug.Print fDialog.SelectedItems(1)
End If
'The msoFileDialogSaveAs dialog does NOT support file filters
'filedialog Filters
'One of the common problems with working with the Application.FileDialog is setting multiple file filters. Below some common examples of how to do this properly. To add a filter for multiple files use the semicolor ;:

Dim fDialog As filedialog
Set fDialog = Application.filedialog(msoFileDialogOpen)
'...
'Optional: Add filters
fDialog.Filters.Clear
fDialog.Filters.Add "All files", "*.*"
fDialog.Filters.Add "Excel files", "*.xlsx;*.xls;*.xlsm"
fDialog.Filters.Add "Text/CSV files", "*.txt;*.csv"
'...
'Be sure to clear your list of filters each time. The FileDialog has its nuisances and often filters are not cleared automatically. Hence, when creating multiple dialogs you might see filters coming from previous executed dialogs if not cleared and re-initiated properly.
