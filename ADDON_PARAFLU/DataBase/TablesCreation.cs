using System;
using System.Runtime.InteropServices;

namespace ADDON_PARAFLU.DataBase
{
    internal class TablesCreation
    {
        private SAPbobsCOM.Company company;
        private SAPbouiCOM.Application application;

        public TablesCreation(SAPbobsCOM.Company company, SAPbouiCOM.Application application)
        {
            this.company = company;
            this.application = application;
        }

        /// <summary>
        /// Creates a table if teh table does not exist.
        /// </summary>
        /// <param name="tabName"> t]Tabel name without "@"</param>
        /// <param name="tabDescription"> Table Description </param>
        /// <param name="tabType"> Table Type. </param>
        public void CreateTable(string tabName, string tabDescription, SAPbobsCOM.BoUTBTableType tabType)
        {
            SAPbobsCOM.UserTablesMD userTablesMD = (SAPbobsCOM.UserTablesMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            try
            {
                if (TableAlreadyExist(tabName))
                    return;

                userTablesMD.TableName = tabName;
                userTablesMD.TableDescription = tabDescription;
                userTablesMD.TableType = tabType;

                if (userTablesMD.Add() != 0)
                {
                    int errorCode = 0;
                    string error = string.Empty;
                    company.GetLastError(out errorCode, out error);
                    throw new Exception($"Erro ao Gerar Tabela: {tabName} erro : [{errorCode} - {error}]");
                }
                else
                {
                    application.StatusBar.SetText("Sucesso ao Gerar Tabela", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao Gerar Tabela: {ex.Message}");
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(userTablesMD);

                userTablesMD = null;
                GC.Collect();
            }
        }

        /// <summary>
        /// verify if the table already exist
        /// </summary>
        /// <param name="table"> nome da tabela. </param>
        /// <returns></returns>
        private bool TableAlreadyExist(string table)
        {
            SAPbobsCOM.Recordset record;
            record = (SAPbobsCOM.Recordset)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = $"SELECT * FROM OUTB WHERE \"TableName\" = '{table}'";
                record.DoQuery(query);

                return record.RecordCount > 0;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao Procurar Tabela: {ex.Message}");
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(record);

                record = null;
                GC.Collect();

            }
        }
    }
}
