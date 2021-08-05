Attribute VB_Name = "test"

Public Sub Synch()

    Dim fsoTarget As FileSystemObject
    Set fsoTarget = New FileSystemObject
    
    Dim fsoFolder As Folder
    Dim fsoSubFld As Folders
    Dim FsoFldCnt As Folder
    Dim intCnt As Integer
    Dim strSFName As String
    Dim strTopFld as String
    

    Set fsoFolder = fsoTarget.GetFolder("C:\Users\sixst\OneDrive")
    
    Set fsoSubFld = fsoFolder.SubFolders
    
        intCnt = 0
    
    For Each FsoFldCnt In fsoSubFld

        strSFName = FsoFldCnt.Name
        Debug.Print strSFName
    
    Next

    strSFName = fsoSubFld.Item(3).Name
    Set fsoFolder = fsoTarget.GetFolder("C:\Users\sixst\OneDrive\" & strSFName)
    
End Sub
