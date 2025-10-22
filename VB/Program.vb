Imports System.IO
Imports System.Text.Json
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace MasterDetailExample

    Friend Module Program

        Sub Main()

            ' Build a path to the JSON data file located in the 'data' folder at the project root.
            Dim dataPath As String = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "data", "nwind_data.json"))

            ' Read the entire JSON file contents into a string.
            Dim json As String = File.ReadAllText(dataPath)

            ' Declare a nullable NWind instance to hold deserialized data.
            Dim nwind As NWindData = Nothing
            Try
                ' Deserialize the JSON payload into an NWind instance using System.Text.Json.
                ' If the JSON shape does not match the NWind type, a JsonException (among others) can be thrown.
                nwind = JsonSerializer.Deserialize(Of NWindData)(json)
            Catch ex As Exception
                ' Log the exception (includes type and message) and abort further processing.
                Console.Error.WriteLine($"Deserialization failed: {ex}")
                Return
            End Try

            ' Create a RichEditDocumentServer instance
            Using wordProcessor As New RichEditDocumentServer()

                ' Build the absolute path to the mail merge template document (DOCX).
                ' The path walks up four directory levels from the base directory to reach the template file.
                Dim templatePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "data", "template.docx")
                If Not File.Exists(templatePath) Then
                    ' Abort if the template is missing, since mail merge cannot proceed without it.
                    Console.Error.WriteLine($"Template not found: {templatePath}")
                    Return
                End If

                ' Load the template into the word processor instance.
                wordProcessor.LoadDocument(templatePath)

                ' Create mail merge options.
                Dim myMergeOptions As MailMergeOptions = wordProcessor.Document.CreateMailMergeOptions()

                ' Assign the data source (customers collection). Null-propagation prevents exceptions if deserialization failed.
                myMergeOptions.DataSource = If(nwind?.Customers, Nothing)

                ' Each merged record starts in a new section, preserving page setup and headers/footers per record.
                myMergeOptions.MergeMode = MergeMode.NewSection
                AddHandler wordProcessor.CalculateDocumentVariable, AddressOf WordProcessor_OnCalculateDocumentVariable

                Dim outputPath As String = Path.Combine(Environment.CurrentDirectory, "result.docx")
                wordProcessor.MailMerge(myMergeOptions, outputPath, DocumentFormat.OpenXml)

                Try
                    Process.Start(New ProcessStartInfo(outputPath) With {.UseShellExecute = True})
                Catch ex As Exception
                    Console.Error.WriteLine($"Unable to open result document: {ex.Message}")
                End Try

            End Using
        End Sub


        Private Sub WordProcessor_OnCalculateDocumentVariable(sender As Object, e As CalculateDocumentVariableEventArgs)
            If e.VariableName = "IMAGE" Then
                Dim pictureProcessor As New RichEditDocumentServer()
                Using imageStream As Stream = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "data", "logo.png"))
                    Dim shape As Shape = pictureProcessor.Document.Shapes.InsertPicture(pictureProcessor.Document.Range.End, DocumentImageSource.FromStream(imageStream))
                    shape.TextWrapping = TextWrappingType.InLineWithText
                    e.Value = pictureProcessor
                    e.Handled = True
                End Using
            End If
        End Sub
    End Module
End Namespace
