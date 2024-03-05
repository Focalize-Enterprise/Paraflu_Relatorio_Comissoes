using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.servicos.Interfaces;
using ADDON_PARAFLU.FORMS.Recursos;
using SAPbouiCOM;
using System.Globalization;
using ADDON_PARAFLU.Uteis.Interfaces;
using SAPbobsCOM;
using ADDON_PARAFLU.Uteis;

namespace ADDON_PARAFLU.FORMS.UserForms
{
    public sealed class EnvioDeRelatorioPorComissoes
    {
        private readonly string formid;
        private readonly SAPbouiCOM.Form form;
        private readonly IAPI _api;
        private readonly IEmail _email;
        private readonly IPDFs _pdfs;
        private readonly DataTable table;
        //Coluna Comissão
        private double totalValue = 0;
        //Coluna Total
        private double totalValue2 = 0;
        private HashSet<int> linhas_selecionadas;

        public EnvioDeRelatorioPorComissoes(IAPI api, IEmail email, IPDFs pdfs)
        {
            _api = api;
            _email = email;
            _pdfs = pdfs;
            _email.GetParamEmail();
            FormCreationParams cp = null;
            string xmlFormCode = Recursos.Recursos.EnvioComissões.ToString();
            linhas_selecionadas = new HashSet<int>();
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
                                        ClearSelection();
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
                                form.Freeze(true);

                                try
                                {
                                    Grid grid = (Grid)form.Items.Item("Item_6").Specific;
                                    // busca a linha mesmo se o grid estiver colapsado
                                    int row = grid.GetDataTableRowIndex(pVal.Row);

                                    // nova versão -- João Copini
                                    if (grid.DataTable.Columns.Item("Selecionado").Cells.Item(row).Value.ToString() == "Y")
                                    {
                                        if(linhas_selecionadas.Contains(row))
                                            return;

                                        //Coluna Comissão
                                        string? valor = table.GetValue("Comissão", row).ToString();
                                        valor ??= "0";
                                        totalValue += Convert.ToDouble(table.GetValue("Comissão", row).ToString(), new CultureInfo("pt-BR"));
                                        double val = Math.Round(totalValue, 2);
                                        ((EditText)form.Items.Item("Item_11").Specific).Value = val.ToString(new CultureInfo("en-US"));

                                        //Coluna Total
                                        string? valor2 = table.GetValue("Total", row).ToString();
                                        valor2 ??= "0";
                                        totalValue2 += Convert.ToDouble(table.GetValue("Total", row).ToString(), new CultureInfo("pt-BR"));
                                        double val2 = Math.Round(totalValue2, 2);
                                        ((EditText)form.Items.Item("Item_9").Specific).Value = val2.ToString(new CultureInfo("en-US"));

                                        linhas_selecionadas.Add(row);
                                    }
                                    else
                                    {
                                        if (!linhas_selecionadas.Contains(row))
                                            return;

                                        //Coluna Comissão
                                        totalValue -= Convert.ToDouble(grid.DataTable.Columns.Item("Comissão").Cells.Item(row).Value.ToString(), new CultureInfo("pt-BR"));
                                        double val = Math.Round(totalValue, 2);
                                        ((EditText)form.Items.Item("Item_11").Specific).Value = val.ToString(new CultureInfo("en-US"));

                                        //Coluna Total
                                        totalValue2 -= Convert.ToDouble(table.GetValue("Total", row).ToString(), new CultureInfo("pt-BR"));
                                        double val2 = Math.Round(totalValue2, 2);
                                        ((EditText)form.Items.Item("Item_9").Specific).Value = val2.ToString(new CultureInfo("en-US"));

                                        linhas_selecionadas.Remove(row);
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
                            if (pVal.ItemUID == "Item_7")
                            {
                                MarcarDesmarcarTodos();
                            }
                        }
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        {

                        }
                        break;
                }
            }
        }

        private void MarcarDesmarcarTodos()
        {
            form.Freeze(true);
            try
            {
                // Y = marcado, N não marcado
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("marcando", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                string marcado = !((SAPbouiCOM.CheckBox)form.Items.Item("Item_7").Specific).Checked ? "Y" : "N";

                DataTable dt = form.DataSources.DataTables.Item("DT_0");
                Grid grid = (Grid)form.Items.Item("Item_6").Specific;


                for (int row = 0; row < table.Rows.Count; row++)
                {
                    if(!table.Columns.Item("Selecionado").Cells.Item(row).Value.ToString()!.Equals(marcado, StringComparison.OrdinalIgnoreCase))
                    {
                        table.Columns.Item("Selecionado").Cells.Item(row).Value = marcado;
                        if (marcado == "Y")
                        {
                            if (linhas_selecionadas.Contains(row))
                                continue;
                            //Coluna Comissão
                            totalValue += Convert.ToDouble(dt.GetValue("Comissão", row).ToString(), new CultureInfo("pt-BR"));

                            //Coluna Total
                            totalValue2 += Convert.ToDouble(dt.GetValue("Total", row).ToString(), new CultureInfo("pt-BR"));

                            linhas_selecionadas.Add(row);
                        }
                        else
                        {
                            totalValue -= Convert.ToDouble(grid.DataTable.Columns.Item("Comissão").Cells.Item(row).Value.ToString(), new CultureInfo("pt-BR"));

                            totalValue2 -= Convert.ToDouble(grid.DataTable.Columns.Item("Total").Cells.Item(row).Value.ToString(), new CultureInfo("pt-BR"));

                            linhas_selecionadas.Remove(row);
                        }
                    }
                }
                totalValue = Math.Round(totalValue, 2);
                ((EditText)form.Items.Item("Item_11").Specific).Value = totalValue.ToString(new CultureInfo("en-US"));

                totalValue2 = Math.Round(totalValue2, 2);
                ((EditText)form.Items.Item("Item_9").Specific).Value = totalValue2.ToString(new CultureInfo("en-US"));

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("marcação finalizada", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            }
            catch (Exception)
            {

            }
            finally
            {
                form.Freeze(false);
            }
        }


        private void AtualizaGrid()
        {
            form.Freeze(true);
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pesquisando dados...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                DataTable dt = form.DataSources.DataTables.Item("DT_0");
                Grid grid = (Grid)form.Items.Item("Item_6").Specific;
                string dataini = ((EditText)this.form.Items.Item("Item_3").Specific).Value;
                string datafim = ((EditText)this.form.Items.Item("Item_4").Specific).Value;
                dataini = dataini.Substring(0, 4) + "/" + dataini.Substring(4, 2) + "/" + dataini.Substring(6, 2);
                datafim = datafim.Substring(0, 4) + "/" + datafim.Substring(4, 2) + "/" + datafim.Substring(6, 2);

                string query = "";
                if (string.IsNullOrEmpty(dataini))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Selecione um perido inicial", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }
                if (string.IsNullOrEmpty(datafim))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Selecione um perido final", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
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
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Dados encontrados...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                form.Freeze(false);
            }
        }
        private (string user, string senha, string past, string crystal) GetDataForBD()
        {
            Recordset recordset = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = @"SELECT ""U_User"", ""U_Pass"", ""U_Past"", ""U_Crystal"" FROM ""@FOC_DB_CONF"" WHERE ""Code"" = '1'";
            recordset.DoQuery(query);
            if (recordset.RecordCount > 0)
            {
                return (recordset.Fields.Item(0).Value.ToString(), recordset.Fields.Item(1).Value.ToString(), recordset.Fields.Item(2).Value.ToString(), recordset.Fields.Item(3).Value.ToString());
            }
            return (string.Empty, string.Empty, string.Empty, string.Empty);
        }

        private void EnviaEmails()
        {
            try
            {
                form.Freeze(true);
                Recordset recordset = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = @"SELECT ""U_Body"" FROM ""@FOC_EMAIL_PARAM"" WHERE ""Code"" = '1'";
                recordset.DoQuery(query);
                (string user, string senha, string past, string crystal) = GetDataForBD();
                bool teste = ((SAPbouiCOM.CheckBox)form.Items.Item("Item_2").Specific).Checked;

                //Vendedores[] values = vendedores.Values.ToArray();

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Enviando {linhas_selecionadas.Count} emails", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                for (int index = 0; index < linhas_selecionadas.Count; index++)
                {
                    //Vendedores vendedores = values[index];
                    // dados do vendedor
                    int row = linhas_selecionadas.ElementAt(index);
                    string cardCode = table.Columns.Item("Código").Cells.Item(row).Value.ToString()!;
                    string email = table.Columns.Item("Email do vendedor").Cells.Item(row).Value.ToString()!;
                    string slpName = table.Columns.Item("Nome do Vendedor").Cells.Item(row).Value.ToString()!;
                    string nomeEmail = email.Split('@').First();

                    //SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"nome: {slpName}, code: {cardCode}, email: {email}, index: {index}");
                    string body = recordset.Fields.Item("U_Body").Value.ToString()!;
                    string reportPath = @$"{System.Windows.Forms.Application.StartupPath}{crystal}";
                    string periodo2 = ((EditText)this.form.Items.Item("Item_4").Specific).Value;
                    string periodo1 = ((EditText)this.form.Items.Item("Item_3").Specific).Value;
                    periodo1 = periodo1.Substring(0, 4) + "-" + periodo1.Substring(4, 2) + "-" + periodo1.Substring(6, 2);
                    periodo2 = periodo2.Substring(0, 4) + "-" + periodo2.Substring(4, 2) + "-" + periodo2.Substring(6, 2);
                    string caminho = $"{past}\\{slpName}_{periodo1}_{periodo2}.pdf";
                    string caminhoPdf = _pdfs.GeraPDF(periodo1, periodo2, cardCode, user, senha, reportPath, caminho);
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Enviando Email...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                    string[] anexos = new string[] { caminhoPdf };
                    _email.EnviarPorEmail(nomeEmail, email, anexos, body, teste);
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

        private void ClearSelection()
        {
            try
            {
                for(int index = 0; index < linhas_selecionadas.Count; index++)
                {
                    int row = linhas_selecionadas.ElementAt(index);
                    table.Columns.Item("Selecionado").Cells.Item(row).Value = "N";
                }

                linhas_selecionadas.Clear();
                totalValue = 0;
                double val = Math.Round(totalValue, 2);
                ((EditText)form.Items.Item("Item_11").Specific).Value = val.ToString(new CultureInfo("en-US"));

                totalValue2 = 0;
                double val2 = Math.Round(totalValue2, 2);
                ((EditText)form.Items.Item("Item_9").Specific).Value = val2.ToString(new CultureInfo("en-US"));
            }
            catch (Exception) { }
        }
    }
}
