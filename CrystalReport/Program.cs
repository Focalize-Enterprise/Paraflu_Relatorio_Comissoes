using CrystalReport.Models;
using System;
using System.Windows.Forms;

namespace CrystalReport
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string [] args)
        {
            string info = args[0];
            args = info.Split(';');
            ReportData data = new ReportData()
            {
                periodo1 = args[0],
                periodo2 = args[1],
                cardCode = args[2],
                server = args[3],
                banco = args[4],
                user = args[5],
                pass = args[6],
                reportExeName = args[7],
                ReportPath = args[8],
            };

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Report(data));
        }
    }
}
