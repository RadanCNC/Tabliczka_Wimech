Public Class RadraftInteropPlugIn
    ' This is a special sub that the Radan application interfaces with.
    Public Sub OnConnectToApplication(app As Radraft.Interop.Application)
        ' Set global variable to the Radan Application that is running
        myApp = app
    End Sub
    ' This is another special sub that interfaces with the Radan GUI.
    ' The menu system that is required in Radan is created here..

    Public Sub OnUpdateGUI()

        ' set the global mac object to our application mac object 
        myMac = myApp.Mac
        ' using myApp, add the top menu with AddMenu 
        If myApp.GUIState = Radraft.Interop.RadGUIState.radPart Then ' Or myApp.GUIState = Radraft.Interop.RadGUIState.radPart 
            ' now add another top menu item and a sub menu item 
            myApp.PluginManager.AddMenu("", "Kreator Tabliczek")
            myApp.PluginManager.AddMenuItem("Kreator Tabliczek", "Kreator Tabliczek", "Click1")

        End If
    End Sub

    Public Sub Click1()

        Dim frm As New Tabliczka_Wimech

        frm.Show()

    End Sub
End Class
