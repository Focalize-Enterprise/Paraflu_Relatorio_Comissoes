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

namespace ADDON_PARAFLU.FORMS.UserForms
{
    public sealed class EnvioDeRelatorioPorComissoes
    {
        private readonly string formid;
        private readonly SAPbouiCOM.Form form;
        private readonly IAPI _api;
        private readonly IEmail _email;
        private DataTable table;
        private double totalValue = 0;
        private Dictionary<string, Vendedores> vendedores = new Dictionary<string, Vendedores>();


        public EnvioDeRelatorioPorComissoes(IAPI api, IEmail email)
        {
            _api = api;
            _email = email;

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
                                case "Item_":
                                    {
                                        //EnviaEmails();
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

                                        double val = 0;
                                        if (vendedores.TryGetValue(vendedor, out Vendedores value))
                                        {
                                            totalValue += Convert.ToDouble(dt.GetValue("Total", row).ToString().Remove(3, 2), new CultureInfo("pt-BR"));
                                            val = Math.Round(totalValue, 2);
                                            ((EditText)form.Items.Item("Item_11").Specific).Value = "R$ " + val.ToString(new CultureInfo("pt-BR"));
                                            return;
                                        }

                                        Vendedores rep = new Vendedores()
                                        {
                                            Code = vendedor,
                                            E_Mail = email,
                                        };

                                        vendedores.Add(vendedor, rep);
                                        totalValue += Convert.ToDouble(dt.GetValue("Total", row).ToString().Remove(0, 3), new CultureInfo("pt-BR"));
                                        val = Math.Round(totalValue, 2);
                                        ((EditText)form.Items.Item("Item_11").Specific).Value = "R$ " + val.ToString(new CultureInfo("pt-BR"));
                                    }
                                    else
                                    {
                                        string vendedor = grid.DataTable.Columns.Item("Código").Cells.Item(row).Value.ToString();
                                        string email = grid.DataTable.Columns.Item("Email do vendedor").Cells.Item(row).Value.ToString();

                                        if (vendedores.TryGetValue(vendedor, out Vendedores value))
                                        {
                                            totalValue -= Convert.ToDouble(grid.DataTable.Columns.Item("Total").Cells.Item(row).Value.ToString().Remove(0, 3), new CultureInfo("pt-BR"));
                                            double val = Math.Round(totalValue, 2);
                                            ((EditText)form.Items.Item("Item_11").Specific).Value = "R$ " + val.ToString(new CultureInfo("pt-BR"));
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

                //string value = ((EditText)this.form.Items.Item("Item_3").Specific).Value;
                //string dataini = ((EditText)this.form.Items.Item("Item_3").Specific).Value;
                //string datafim = ((EditText)this.form.Items.Item("Item_4").Specific).Value;
                //dataini = dataini.Substring(0, 4) + "/" + dataini.Substring(4, 2) + "/" + dataini.Substring(6, 2);
                //datafim = datafim.Substring(0, 4) + "/" + datafim.Substring(4, 2) + "/" + datafim.Substring(6, 2);
                //if (string.IsNullOrEmpty(value))
                //{
                //    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Selecione um perido", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                //    return;
                //}
                dt.ExecuteQuery(Queries.Notas_Fiscais);
                grid.Columns.Item("Selecionado").Type = BoGridColumnType.gct_CheckBox;
                //string query = Queries.Notas_Fiscais.Replace("Dataini", dataini).Replace("Datafim", datafim);
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

        //private void EnviaEmails()
        //{
        //    form.Freeze(true);
        //    DataTable dt = form.DataSources.DataTables.Item("DT_0");
        //    try
        //    {
        //        form.Freeze(true);
        //        Vendedores[] values = vendedores.Values.ToArray();
        //        for (int index = 0; index < values.Length; index++)
        //        {
        //            Vendedores vendedores = values[index];

        //            string cardCode = vendedores.Code;
        //            string serial = vendedores.NfsAsSQLValue();
        //            string periodo1 = ((EditText)this.form.Items.Item("Item_5").Specific).Value;
        //            string periodo2 = ((EditText)this.form.Items.Item("Item_4").Specific).Value;
        //            periodo1 = periodo1.Substring(0, 4) + "-" + periodo1.Substring(4, 2) + "-" + periodo1.Substring(6, 2);
        //            periodo2 = periodo2.Substring(0, 4) + "-" + periodo2.Substring(4, 2) + "-" + periodo2.Substring(6, 2);
        //            string body = ((SAPbouiCOM.EditText)form.Items.Item("ETTX_EM").Specific).Value;
        //            string[] anexos = new string[2] { caminhoPdf, caminhoExcel };

        //            _email.EnviarPorEmail(value.Split('@').First(), value, anexos, body);

        //            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Email enviado com sucesso!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Erro ao enviar email:{ex.Message}");
        //    }
        //    finally
        //    {
        //        form.Freeze(false);
        //    }
        //}
    }
}
