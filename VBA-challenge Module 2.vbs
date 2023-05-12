Attribute VB_Name = "Module2"
Sub Mod_2_Assign_Loop()

    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
        Call Module1.Mod_2_Assign
    Next

End Sub
