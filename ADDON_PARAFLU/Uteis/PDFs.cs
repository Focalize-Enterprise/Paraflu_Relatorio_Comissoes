using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.Uteis.Interfaces;
using CristalReportsLibrary.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADDON_PARAFLU.Uteis
{
    public class PDFs : IPDFs
    {
        private readonly IAPI _api;
        internal string periodo1 { get; set; }
        internal string periodo2 { get; set; }
        internal string cardCode { get; set; }
        internal string DBuser { get; set; }
        internal string DBsenha { get; set; }

        public PDFs(IAPI api)
        {
            _api = api;
        }
        public string GeraPDF(string periodo1, string periodo2, string cardCode, string DBuser, string DBsenha, string caminho = "")
        {
            string server = _api.Company.Server;
            string banco = _api.Company.CompanyDB;
            string user = DBuser;
            string pass = DBsenha;
            ReportData data = new ReportData()
            {
                banco = banco,
                cardCode = cardCode,
                pass = pass,
                user = user,
                periodo1 = periodo1,
                periodo2 = periodo2,
                server = server,
                reportExeName = "Report",
            };

            return CristalReportsLibrary.Report.GenerateReport(data, caminho);
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
