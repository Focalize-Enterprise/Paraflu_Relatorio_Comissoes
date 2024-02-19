using SAPbobsCOM;
using SAPbouiCOM;
using System;

namespace ADDON_PARAFLU.DataBase
{
    public class SetUp
    {
        /// <summary>
        /// start the script for the DataBase fields and Tables creation
        /// </summary>
        /// <param name="company"> DI Company </param>
        public static bool StartSetUp(SAPbobsCOM.Company company)
        {
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Validando Tabelas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                SetupTables(SAPbouiCOM.Framework.Application.SBO_Application, company);

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Validando campos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                SetupFields(SAPbouiCOM.Framework.Application.SBO_Application, company);

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Sucesso ao Validar banco", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                return true;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        /// <summary>
        /// Creates all tables for the add-on
        /// </summary>
        /// <param name="application"> UI application control </param>
        /// <param name="company"> DI Company </param>
        private static void SetupTables(SAPbouiCOM.Application application, SAPbobsCOM.Company company)
        {
            //User Tables
            TablesCreation tables = new TablesCreation(company, application);
            //Valores para a DataBase
            tables.CreateTable("FOC_DB_CONF", "FOC_DB_CONF", BoUTBTableType.bott_NoObject);
            //Valores para envio de Email
            tables.CreateTable("FOC_EMAIL_PARAM", "FOC_EMAIL_PARAM", BoUTBTableType.bott_NoObject);
            
        }

        /// <summary>
        /// Set the tables fields.
        /// </summary>
        /// <param name="application"> UI application control </param>
        /// <param name="company"> DI Company </param>
        private static void SetupFields(SAPbouiCOM.Application application, SAPbobsCOM.Company company)
        {
            var valids = new DataBase.ValidValue[2]
            {
                new DataBase.ValidValue() { Value = "Y", Description = "Sim" },
                new DataBase.ValidValue() { Value = "N", Description = "Não" },
            };
            TableFields fields = new TableFields(company, application);
            // user tables
            //
            //Criacao de campos na tabela de email
            //
            fields.CreateFields("FOC_EMAIL_PARAM", "Email", "Email", BoFieldTypes.db_Alpha, 100, string.Empty);
            fields.CreateFields("FOC_EMAIL_PARAM", "senha", "senha", BoFieldTypes.db_Memo, 254, string.Empty);
            fields.CreateFields("FOC_EMAIL_PARAM", "host", "host", BoFieldTypes.db_Alpha, 100, string.Empty);
            fields.CreateFields("FOC_EMAIL_PARAM", "porta", "porta", BoFieldTypes.db_Numeric, 10, string.Empty);
            fields.CreateFields("FOC_EMAIL_PARAM", "Body", "Corpo do Email", BoFieldTypes.db_Memo, 254, string.Empty);
            //
            // campos da configuracao do banco/pasta para salvar PDF
            //
            fields.CreateFields("FOC_DB_CONF", "User", "Usuario", BoFieldTypes.db_Memo, 254, string.Empty);
            fields.CreateFields("FOC_DB_CONF", "Pass", "Senha", BoFieldTypes.db_Memo, 254, string.Empty);
            fields.CreateFields("FOC_DB_CONF", "Past", "Pasta para Salvar PDF", BoFieldTypes.db_Memo, 254, string.Empty);
            fields.CreateFields("FOC_DB_CONF", "Item", "Item Comissão", BoFieldTypes.db_Alpha, 254, string.Empty);
            fields.CreateFields("FOC_DB_CONF", "Util", "Utilização nota", BoFieldTypes.db_Numeric, 10, string.Empty);
            //
            //Campos na tabela OINV
            fields.CreateFields("OINV", "FOC_StatusComissao", "Status Comissão", BoFieldTypes.db_Alpha, 1, string.Empty, validvalues: valids, systemTable: true);
            fields.CreateFields("OINV", "FOC_ValorComissao", "Valor da Comissão (add-on)", BoFieldTypes.db_Float, 10, string.Empty, BoFldSubTypes.st_Price, systemTable: true);
            fields.CreateFields("OINV", "FOC_DocEntryNFCom", "DocEntry da NF de Comissão", BoFieldTypes.db_Numeric, 10, string.Empty, systemTable: true);

        }

        /// <summary>
        /// ciração dos UDOS
        /// </summary>
        /// <param name="application"> controle UI do SAP </param>
        /// <param name="company"> Controle da di api </param>

    }
}
