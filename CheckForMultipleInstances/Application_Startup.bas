Attribute VB_Name = "Application_Startup"
Option Compare Database
Option Explicit

Public Function StartApp()
On Error GoTo CATCH_ 'TRY_

    Dim accessApp As Access.Application
    Set accessApp = VBA.GetObject(CurrentProject.FullName)
    
    Dim commandParameter As String
        commandParameter = VBA.CStr(VBA.command)
            
    If accessApp.hWndAccessApp = Application.hWndAccessApp Then
        'Es gibt nur eine Instanz der Anwendung
        Call Application_Startup.StartAppForm(commandParameter)
                
    Else
        'Es wurde eine zweite Instanz der Anwendung gestartet
        accessApp.Run "StartAppForm", commandParameter
        
        'die zweite Instanz wieder schließen
        Application.Quit acQuitSaveNone
                
    End If
        
    GoTo FINALLY_
        
CATCH_:
    Debug.Print "Error: " & Err.Description
    Resume FINALLY_
    
FINALLY_:
    On Error Resume Next
    Set accessApp = Nothing
    Exit Function
End Function


Public Sub StartAppForm(ByVal commandParameter As String)
On Error GoTo CATCH_ 'TRY_
    
    If VBA.Len(commandParameter) > 0 Then
       
        DoCmd.Close acForm, "AppForm", acSaveNo
        DoCmd.OpenForm "AppForm", , , , , , commandParameter
        
    End If

    GoTo FINALLY_

CATCH_:
    Debug.Print "Error in StartAppForm: " & Err.Description
    Resume FINALLY_
    
FINALLY_:
    On Error Resume Next
    Exit Sub
End Sub
