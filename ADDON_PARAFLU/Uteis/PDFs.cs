using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.Uteis.Interfaces;
using SAPbobsCOM;
using System.Diagnostics;
using System.IO;

namespace ADDON_PARAFLU.Uteis
{
    public class PDFs : IPDFs
    {
        private readonly IAPI _api;
        private bool timeout = false;

        public PDFs(IAPI api)
        {
            _api = api;
        }
        public string GeraPDF(string periodo1, string periodo2, string cardCode, string DBuser, string DBsenha, string reportPath, string pdfPath)
        {
            string server = _api.Company!.Server;
            string banco = _api.Company.CompanyDB;
            string user = DBuser;
            string pass = DBsenha;
            string servicePath = @$"{Application.StartupPath}\\Report\\CrystalReport.exe";
            string exeArgs = $@"""{periodo1};{periodo2};{cardCode};{server};{banco};{user};{pass};{reportPath};{pdfPath}""";
            ProcessStartInfo info = new(servicePath)
            {
                Arguments = exeArgs
            };

            var process = Process.Start(info);
            Stopwatch watch = new Stopwatch();
            bool sucess = false;
            if (process is not null)
            {
                watch.Start();
                while (!process.HasExited && watch.Elapsed.Seconds < 20) { }
                if (process.ExitCode == 0)
                    sucess = true;
                else if (!process.HasExited)
                    process.Close();

                watch.Stop();
            }
            sucess = File.Exists(pdfPath);

            return sucess ? pdfPath : "";
        }

        private (string user, string senha) GetDataForBD()
        {
            Recordset recordset = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = @"SELECT U_User, U_Pass FROM ""@FOC_DB_CONF"" WHERE Code = '1'";
            recordset.DoQuery(query);
            if (recordset.RecordCount > 0)
            {
                //return (Security.Decrypt(recordset.Fields.Item(0).Value.ToString()), Security.Decrypt(recordset.Fields.Item(1).Value.ToString()));
            }

            return (string.Empty, string.Empty);
        }
    }
}
