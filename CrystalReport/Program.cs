using CrystalDecisions.CrystalReports.Engine;
using CrystalReport.Models;
using System;
using System.Windows.Forms;
using System.Xml.Linq;

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

            GenerateReport(data);
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Report(data));
        }

        static void GenerateReport(ReportData _data)
        {
            // lembrar de mudar o Report para o servidor da Pasinato caso haja um problema de vendedor da base 
            ReportDocument report = new ReportDocument();
            // load do crystal passando o caminho da aplicação.
            report.Load(_data.reportExeName);
            report.Refresh();
            report.DataSourceConnections.Clear();

            // parâmetros do crystal. 
            report.SetParameterValue("DateI", _data.periodo1);
            report.SetParameterValue("DateF", _data.periodo2);
            report.SetParameterValue("SlpCode", _data.cardCode);

            for (int i = 0; i < report.Subreports.Count; i++)
            {
                report.SetParameterValue("DateI", _data.periodo1, report.Subreports[i].Name);
                report.SetParameterValue("DateF", _data.periodo2, report.Subreports[i].Name);
                report.SetParameterValue("SlpCode", _data.cardCode);
            }

            // dados de conexão OLE DB SQL
            for (int index = 0; index < report.DataSourceConnections.Count; index++)
            {
                report.DataSourceConnections[index].SetConnection(_data.server, _data.banco, _data.user, _data.pass);
            }
            report.SetDatabaseLogon(_data.user, _data.pass);
            // exporta para o PDF
            string fileName = _data.ReportPath;

            report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, fileName);
            report.Close();
            report.Dispose();
        }
    }
}
