using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace MasterDetailExample
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(@"result.docx") {
                UseShellExecute = true
            };
            p.Start();
        }
    }
}
