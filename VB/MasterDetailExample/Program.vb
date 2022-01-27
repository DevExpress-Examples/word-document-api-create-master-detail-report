Imports System
Imports System.Diagnostics

Namespace MasterDetailExample

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            Dim mProcessor As MergeProcessor = New MergeProcessor()
            mProcessor.Start()
            Call Process.Start("result.docx")
        End Sub
    End Module
End Namespace
