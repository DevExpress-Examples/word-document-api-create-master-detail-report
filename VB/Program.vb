Imports System
Imports System.Diagnostics
Imports DevExpress.DataAccess.Native.EntityFramework

Namespace MasterDetailExample

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        Sub Main()
            Dim mProcessor As MergeProcessor = New MergeProcessor()
            mProcessor.Start()
            Dim processor As New Process()
            processor.StartInfo = New ProcessStartInfo("result.docx") With {.UseShellExecute = True}
            processor.Start()
        End Sub
    End Module
End Namespace
