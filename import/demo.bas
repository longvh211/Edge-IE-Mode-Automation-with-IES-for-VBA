Attribute VB_Name = "Demo"
Sub TestEdgeIeMode()
'----------------------------------------------------------------------------------
' Note : Make sure that the target webpage is already loaded in
'        Edge Ie Mode before running this.
' Guide: 1. First run "taskkill /f /im msedge.exe" in cmd to close all
'           hidden edge instances.
'        2. Load this url "https://www.hsbc.com.sg/security/" with Edge IE Mode.
'        3. Execute this procedure. If successful, it will attempt to input into
'           the username field.
'----------------------------------------------------------------------------------

   'Verify if target window is found
    titleToFind = "Username | Log on | HSBC"
    Set ieDoc = GetEdgeIeDOM(titleToFind)
    If ieDoc Is Nothing Then
        MsgBox "The webpage cannot be found on Edge IE Mode! Has it been loaded under Edge IE Mode?"
        End
    End If

   'If found, perform automation
    ieDoc.getElementById("username").Value = "Test Successful!!"

End Sub
