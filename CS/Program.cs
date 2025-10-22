using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using MasterDetailExample;
using System.Diagnostics;
using System.Text.Json;

// Build a path to the JSON data file located in the 'data' folder at the project root.
string dataPath = Path.GetFullPath(
    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "data", "nwind_data.json"));

// Read the entire JSON file contents into a string.
string json = File.ReadAllText(dataPath);

// Declare a nullable NWind instance to hold deserialized data.
NWindData? nwind;
try
{
    // Deserialize the JSON payload into an NWind instance using System.Text.Json.
    // If the JSON shape does not match the NWind type, a JsonException (among others) can be thrown.
    nwind = JsonSerializer.Deserialize<NWindData>(json);
}
catch (Exception ex)
{
    // Log the exception (includes type and message) and abort further processing.
    Console.Error.WriteLine($"Deserialization failed: {ex}");
    return;
}

// Create a RichEditDocumentServer instance
using var wordProcessor = new RichEditDocumentServer();

// Build the absolute path to the mail merge template document (DOCX).
// The path walks up four directory levels from the base directory to reach the template file.
var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "data", "template.docx");
if (!File.Exists(templatePath))
{
    // Abort if the template is missing, since mail merge cannot proceed without it.
    Console.Error.WriteLine($"Template not found: {templatePath}");
    return;
}
// Load the template into the word processor instance.
wordProcessor.LoadDocument(templatePath);

// Create mail merge options.
MailMergeOptions myMergeOptions = wordProcessor.Document.CreateMailMergeOptions();

// Assign the data source (customers collection). Null-propagation prevents exceptions if deserialization failed.
myMergeOptions.DataSource = nwind?.Customers;

// Each merged record starts in a new section, preserving page setup and headers/footers per record.
myMergeOptions.MergeMode = MergeMode.NewSection;
wordProcessor.CalculateDocumentVariable += WordProcessor_OnCalculateDocumentVariable;

var outputPath = Path.Combine(Environment.CurrentDirectory, "result.docx");
wordProcessor.MailMerge(myMergeOptions, outputPath, DocumentFormat.OpenXml);

try
{
    Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Unable to open result document: {ex.Message}");
}

void WordProcessor_OnCalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
{
    if (e.VariableName == "IMAGE")
    {
        var pictureProcessor = new RichEditDocumentServer();
        using var imageStream = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "data", "logo.png"));
        Shape shape = pictureProcessor.Document.Shapes.InsertPicture(pictureProcessor.Document.Range.End, DocumentImageSource.FromStream(imageStream));
        shape.TextWrapping = TextWrappingType.InLineWithText;
        e.Value = pictureProcessor;
        e.Handled = true;
    }
}





