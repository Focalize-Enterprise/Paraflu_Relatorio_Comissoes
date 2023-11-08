using CrystalDecisions.CrystalReports.Engine;
using CrystalReport.Models;
using System;
using System.Windows.Forms;

namespace CrystalReport
{
    public partial class Report : Form
    {
        private readonly ReportData _data;

        public Report(ReportData data)
        {
            _data = data;
            InitializeComponent();
        }

        private void GenerateReport()
        {
            // lembrar de mudar o Report para o servidor da Pasinato caso haja um problema de vendedor da base 
            ReportDocument report = new ReportDocument();
            // load do crystal passando o caminho da aplicação.
            string path = AppDomain.CurrentDomain.BaseDirectory;
            report.Load(path + $"\\{_data.reportExeName}.rpt");
            report.Refresh();
            report.DataSourceConnections.Clear();

            // parâmetros do crystal. 
            report.SetParameterValue("dataInicio", _data.periodo1);
            report.SetParameterValue("dataFim", _data.periodo2);
            report.SetParameterValue("SlpCode", _data.cardCode);

            for (int i = 0; i < report.Subreports.Count; i++)
            {
                report.SetParameterValue("dataInicio", _data.periodo1, report.Subreports[i].Name);
                report.SetParameterValue("dataFim", _data.periodo2, report.Subreports[i].Name);
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

        private void Report_Load(object sender, EventArgs e)
        {
            try
            {
                GenerateReport();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.InnerException);
            }
        }
    }
}
