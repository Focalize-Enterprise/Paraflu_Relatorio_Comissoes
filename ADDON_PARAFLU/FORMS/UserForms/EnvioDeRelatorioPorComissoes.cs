using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.servicos.Interfaces;
using ADDON_PARAFLU.FORMS.Recursos;
using Microsoft.Extensions.DependencyInjection;
using SAPbouiCOM;
using ADDON_PARAFLU.Uteis;
using System.Globalization;
using ADDON_PARAFLU.Uteis.Interfaces;
using SAPbobsCOM;

namespace ADDON_PARAFLU.FORMS.UserForms
{
    public sealed class EnvioDeRelatorioPorComissoes
    {
        private readonly string formid;
        private readonly SAPbouiCOM.Form form;
        private readonly IAPI _api;
        private readonly IEmail _email;
        private readonly IPDFs _pdfs;
        private DataTable table;
        private double totalValue = 0;
        private Dictionary<string, Vendedores> vendedores = new Dictionary<string, Vendedores>();


        public EnvioDeRelatorioPorComissoes(IAPI api, IEmail email, IPDFs pdfs)
        {
            _api = api;
            _email = email;
            _pdfs = pdfs;

            FormCreationParams cp = null;
            string xmlFormCode = Recursos.Recursos.EnvioComissões.ToString();
            try
            {
                cp = ((FormCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams)));
                cp.FormType = "EnvioComissoes_Form";
                cp.XmlData = xmlFormCode;
                cp.UniqueID = "EnvioComissoes_Form";
                this.form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.AddEx(cp);
                this.formid = this.form.UniqueID;
                CustomInitialize();
                table = form.DataSources.DataTables.Item("DT_0");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao crair o formulário: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (form != null)
                {
                    form.Visible = true;
                }
            }
        }
        private void CustomInitialize()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }
        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (FormUID != formid)
                return;
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        {
                            switch (pVal.ItemUID)
                            {
                                case "Item_1":
                                    {
                                        AtualizaGrid();
                                    }
                                    break;
                                case "Item_8":
                                    {
                                        EnviaEmails();
                                    }
                                    break;
                            }
                        }
                        break;
                }
            }
            else
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        {
                            if (pVal.ItemUID == "Item_6" && pVal.ColUID == "Selecionado" && !pVal.BeforeAction)
                            {
                                try
                                {
                                    form.Freeze(true);
                                    Grid grid = (Grid)form.Items.Item("Item_6").Specific;
                                    int row = grid.GetDataTableRowIndex(pVal.Row);
                                    DataTable dt = form.DataSources.DataTables.Item("DT_0");

                                    if (grid.DataTable.Columns.Item("Selecionado").Cells.Item(row).Value.ToString() == "Y")
                                    {
                                        string vendedor = grid.DataTable.Columns.Item("Código").Cells.Item(row).Value.ToString();
                                        string email = grid.DataTable.Columns.Item("Email do vendedor").Cells.Item(row).Value.ToString();

                                        string valor = dt.GetValue("Comissão", row).ToString();
                                        double val = 0;
                                        if (vendedores.TryGetValue(vendedor, out Vendedores value))
                                        {
                                            totalValue += Convert.ToDouble(dt.GetValue("Comissão", row).ToString(), new CultureInfo("pt-BR"));
                                            val = Math.Round(totalValue, 2);
                                            ((EditText)form.Items.Item("Item_11").Specific).Value = val.ToString(new CultureInfo("en-US"));
                                            return;
                                        }

                                        Vendedores rep = new Vendedores()
                                        {
                                            Code = vendedor,
                                            E_Mail = email,
                                        };

                                        vendedores.Add(vendedor, rep);
                                        totalValue += Convert.ToDouble(dt.GetValue("Comissão", row).ToString(), new CultureInfo("pt-BR"));
                                        val = Math.Round(totalValue, 2);
                                        ((EditText)form.Items.Item("Item_11").Specific).Value = val.ToString(new CultureInfo("en-US"));
                                    }
                                    else
                                    {
                                        string vendedor = grid.DataTable.Columns.Item("Código").Cells.Item(row).Value.ToString();
                                        string email = grid.DataTable.Columns.Item("Email do vendedor").Cells.Item(row).Value.ToString();

                                        if (vendedores.TryGetValue(vendedor, out Vendedores value))
                                        {
                                            totalValue -= Convert.ToDouble(grid.DataTable.Columns.Item("Comissão").Cells.Item(row).Value.ToString(), new CultureInfo("pt-BR"));
                                            double val = Math.Round(totalValue, 2);
                                            ((EditText)form.Items.Item("Item_11").Specific).Value = val.ToString(new CultureInfo("en-US"));
                                            return;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"erro ao atualizar o valor de cotação {ex.Message}");
                                }
                                finally
                                {
                                    form.Freeze(false);
                                }
                            }
                        }
                        break;
                }
            }
        }
        private void AtualizaGrid()
        {
            form.Freeze(true);
            try
            {
                DataTable dt = form.DataSources.DataTables.Item("DT_0");
                Grid grid = (Grid)form.Items.Item("Item_6").Specific;
                string dataini = ((EditText)this.form.Items.Item("Item_3").Specific).Value;
                string datafim = ((EditText)this.form.Items.Item("Item_4").Specific).Value;
                dataini = dataini.Substring(0, 4) + "/" + dataini.Substring(4, 2) + "/" + dataini.Substring(6, 2);
                datafim = datafim.Substring(0, 4) + "/" + datafim.Substring(4, 2) + "/" + datafim.Substring(6, 2);
                string query = "";
                if (string.IsNullOrEmpty(dataini))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Selecione um perido", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }
                if (string.IsNullOrEmpty(datafim))
                {
                }
                if (_api.Company.DbServerType != BoDataServerTypes.dst_HANADB)
                    query = Queries.Notas_Fiscais_SQL.Replace("Dataini", dataini).Replace("Datafim", datafim);
                else
                    query = Queries.Notas_Fiscais_HANA.Replace("Dataini", dataini).Replace("Datafim", datafim);
                dt.ExecuteQuery(query);
                grid.Columns.Item("Selecionado").Type = BoGridColumnType.gct_CheckBox;
                for (int index = 0; index < grid.Columns.Count; index++)
                    grid.Columns.Item(index).Editable = false;
                grid.Columns.Item("Selecionado").Editable = true;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao buscar os dados: {ex.Message}");
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void EnviaEmails()

        {
            form.Freeze(true);
            DataTable dt = form.DataSources.DataTables.Item("DT_0");
            try
            {
                form.Freeze(true);
                Vendedores[] values = vendedores.Values.ToArray();
                for (int index = 0; index < values.Length; index++)
                {
                    Vendedores vendedores = values[index];
                    string reportPath = @$"{System.Windows.Forms.Application.StartupPath}ReportComissões.rpt";
                    string user = "sa";
                    string caminho = "";
                    string senha = "@B1Admin123#";
                    string cardCode = vendedores.Code;
                    string periodo2 = ((EditText)this.form.Items.Item("Item_4").Specific).Value;
                    string periodo1 = ((EditText)this.form.Items.Item("Item_3").Specific).Value;
                    periodo1 = periodo1.Substring(0, 4) + "-" + periodo1.Substring(4, 2) + "-" + periodo1.Substring(6, 2);
                    periodo2 = periodo2.Substring(0, 4) + "-" + periodo2.Substring(4, 2) + "-" + periodo2.Substring(6, 2);
                    if(string.IsNullOrEmpty(caminho))
                    caminho = $"C:\\Temp\\{cardCode}.pdf";
                    //string body = ((SAPbouiCOM.EditText)form.Items.Item("ETTX_EM").Specific).Value;
                    string caminhoPdf = _pdfs.GeraPDF(periodo1, periodo2, cardCode, user, senha, reportPath, caminho);
                    string[] anexos = new string[] { caminhoPdf };

                    _email.EnviarPorEmail(vendedores.E_Mail.Split('@').First(), vendedores.E_Mail, anexos);

                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Email enviado com sucesso!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Erro ao enviar email:{ex.Message}");
            }
            finally
            {
                form.Freeze(false);
            }
        }
    }
}
