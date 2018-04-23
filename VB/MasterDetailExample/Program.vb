Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Windows.Forms

Namespace MasterDetailExample
	Friend NotInheritable Class Program
		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		Private Sub New()
		End Sub
		<STAThread> _
		Shared Sub Main()
			Dim mProcessor As New MergeProcessor()
			mProcessor.Start()
			Process.Start("result.docx")
		End Sub
	End Class
End Namespace
