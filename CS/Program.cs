using System.Diagnostics;

namespace MasterDetailExample
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main() {
            MergeProcessor mProcessor = new MergeProcessor();
            mProcessor.Start();
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(@"result.docx") {
                UseShellExecute = true
            };
            p.Start();
        }
    }
}
