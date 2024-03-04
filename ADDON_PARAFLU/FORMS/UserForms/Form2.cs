using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.FORMS.Recursos;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ADDON_PARAFLU.FORMS.UserForms
{
    public sealed class Form2
    {
        private readonly DBDataSource _dataSource;
        private readonly IAPI _api;
        private readonly string formid;
        private readonly SAPbouiCOM.Form form;
        private DataTable table;
        private double totalValue = 0;
        private double totalValue2 = 0;
        private Dictionary<int, Nota> nota = new Dictionary<int, Nota>();
        private HashSet<int> notas = new HashSet<int>();


        public Form2(IAPI api)
        {
            _api = api;

            FormCreationParams cp = null;
            string xmlFormCode = Recursos.Recursos.Form2.ToString();
            try
            {
                cp = ((FormCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams)));
                cp.FormType = "Form2";
                cp.XmlData = xmlFormCode;
                cp.UniqueID = "Form2";
                this.form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.AddEx(cp);
                this.formid = this.form.UniqueID;
                table = form.DataSources.DataTables.Item("DT_0");
                CustomInitialize();
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
            InformaçõesPadrão();
        }

        private void InformaçõesPadrão()
        {
            ((EditText)this.form.Items.Item("Item_16").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            ((EditText)this.form.Items.Item("Item_20").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            ((EditText)this.form.Items.Item("Item_21").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
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
                                case "Item_3":
                                    {
                                        Limpar();
                                    }
                                    break;
                                case "Item_8":
                                    {
                                        Listar();
                                    }
                                    break;
                                case "Item_23":
                                    {
                                        GravarEsboço();
                                    }
                                    break;
                                case "Item_10":
                                    {
                                        MarcarDesmarcarTodos();
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
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        {
                            form.Freeze(true);
                            try
                            {
                                string value = string.Empty;
                                IChooseFromListEvent CFLEvento = (IChooseFromListEvent)pVal;
                                DataTable dataTable = CFLEvento.SelectedObjects;

                                if (dataTable != null)
                                    value = dataTable.GetValue(0, 0).ToString();
                                if (string.IsNullOrEmpty(value) || dataTable is null || value == ((EditText)this.form.Items.Item("Item_1").Specific).Value)
                                    return;

                                switch (CFLEvento.ChooseFromListUID)
                                {
                                    case "CFL_0":
                                        {
                                            ((EditText)this.form.Items.Item("Item_2").Specific).Value = dataTable.GetValue("SlpName", 0).ToString();
                                            ((EditText)this.form.Items.Item("Item_1").Specific).Value = value;
                                        }
                                        break;
                                }
                            }
                            catch (Exception ex)
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"erro ao selecionar Vendedor {ex.Message}");
                            }
                            finally
                            {
                                form.Freeze(false);
                            }
                        }
                        break;
                    case BoEventTypes.et_CLICK:
                        {
                            if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Check" && !pVal.BeforeAction)
                            {
                                try
                                {
                                    form.Freeze(true);
                                    Grid grid = (Grid)form.Items.Item("Item_9").Specific;
                                    int row = grid.GetDataTableRowIndex(pVal.Row);
                                    DataTable dt = form.DataSources.DataTables.Item("DT_0");

                                    if (grid.DataTable.Columns.Item("Check").Cells.Item(row).Value.ToString() == "Y")
                                    {
                                        int codNota = int.Parse(grid.DataTable.Columns.Item("DocEntry").Cells.Item(row).Value.ToString());
                                        double val = 0;
                                        double val2 = 0;

                                        if (notas.TryGetValue(codNota, out _))
                                        {
                                            //Coluna Comissão
                                            totalValue += Convert.ToDouble(dt.GetValue("Comissao", row).ToString(), new CultureInfo("pt-BR"));
                                            val = Math.Round(totalValue, 2);
                                            ((EditText)form.Items.Item("Item_12").Specific).Value = val.ToString(new CultureInfo("en-US"));

                                            //Coluna Total
                                            totalValue2 += Convert.ToDouble(dt.GetValue("Total", row).ToString(), new CultureInfo("pt-BR"));
                                            val2 = Math.Round(totalValue2, 2);
                                            ((EditText)form.Items.Item("Item_26").Specific).Value = val2.ToString(new CultureInfo("en-US"));

                                            return;
                                        }
                                        //Coluna Comissão
                                        notas.Add(codNota);
                                        totalValue += Convert.ToDouble(dt.GetValue("Comissao", row).ToString(), new CultureInfo("pt-BR"));
                                        val = Math.Round(totalValue, 2);
                                        ((EditText)form.Items.Item("Item_12").Specific).Value = val.ToString(new CultureInfo("en-US"));

                                        //Coluna Total
                                        totalValue2 += Convert.ToDouble(dt.GetValue("Total", row).ToString(), new CultureInfo("pt-BR"));
                                        val2 = Math.Round(totalValue2, 2);
                                        ((EditText)form.Items.Item("Item_26").Specific).Value = val2.ToString(new CultureInfo("en-US"));
                                    }
                                    else
                                    {
                                        int codNota = int.Parse(grid.DataTable.Columns.Item("DocEntry").Cells.Item(row).Value.ToString());
                                        if (notas.TryGetValue(codNota, out _))
                                        {
                                            totalValue -= Convert.ToDouble(grid.DataTable.Columns.Item("Comissao").Cells.Item(row).Value.ToString(), new CultureInfo("pt-BR"));
                                            double val = Math.Round(totalValue, 2);
                                            ((EditText)form.Items.Item("Item_12").Specific).Value = val.ToString(new CultureInfo("en-US"));

                                            totalValue2 -= Convert.ToDouble(grid.DataTable.Columns.Item("Total").Cells.Item(row).Value.ToString(), new CultureInfo("pt-BR"));
                                            double val2 = Math.Round(totalValue2, 2);
                                            ((EditText)form.Items.Item("Item_26").Specific).Value = val2.ToString(new CultureInfo("en-US"));

                                            notas.Remove(codNota);
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
        private void GravarEsboço()
        {
            //Dados para gerar doc
            Recordset recordset = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = @"SELECT ""U_User"", ""U_Pass"", ""U_Past"", ""U_Item"", ""U_Util"" FROM ""@FOC_DB_CONF"" WHERE ""Code"" = '1'";
            recordset.DoQuery(query);
            Grid grid = (Grid)form.Items.Item("Item_9").Specific;
            int row = 0;

            string dtLanç = ((EditText)this.form.Items.Item("Item_20").Specific).Value;
            string dtVenc = ((EditText)this.form.Items.Item("Item_21").Specific).Value;
            string dtEmi = ((EditText)this.form.Items.Item("Item_16").Specific).Value;
            string serieNF = ((EditText)form.Items.Item("Item_18").Specific).Value;
            string numNF = ((EditText)form.Items.Item("Item_17").Specific).Value;
            double comissao = Convert.ToDouble(((EditText)form.Items.Item("Item_12").Specific).Value, new CultureInfo("en-US"));
            string codParc = grid.DataTable.Columns.Item("U_FOC_CodPN").Cells.Item(row).Value.ToString();
            string itemParam = recordset.Fields.Item(3).Value.ToString();
            string util = recordset.Fields.Item(4).Value.ToString();
            string sequenceModel = "46";

            if (string.IsNullOrEmpty(dtLanç))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Necessário preencher Data de Lançamento da Nota");
                return;
            }
            if (string.IsNullOrEmpty(dtVenc))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Necessário preencher Data de Vencimento da Nota");
                return;
            }
            if (string.IsNullOrEmpty(dtEmi))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Necessário preencher Data de Emissão da Nota");
                return;
            }
            if (string.IsNullOrEmpty(serieNF))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Necessário preencher Número de Série da Nota");
                return;
            }
            if (string.IsNullOrEmpty(numNF.ToString()))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Necessário preencher Número da Nota da Nota");
                return;
            }
            GerarDoc(dtLanç, dtVenc, dtEmi, serieNF, numNF, comissao, codParc, itemParam, util, sequenceModel);
            Listar();
        }

        private void GerarDoc(string dtLanç, string dtVenc, string dtEmi, string serieNF, string numNF, double comissao, string codParc, string itemParam, string util, string sequenceModel)
        {
            form.Freeze(true);
            bool checkRet = ((SAPbouiCOM.CheckBox)form.Items.Item("Item_24").Specific).Checked;
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Gravando Esboço...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                Documents documents = (Documents)_api.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                documents.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                documents.DocDate = DateTime.ParseExact(dtLanç, "yyyyMMdd", CultureInfo.InvariantCulture);
                documents.TaxDate = DateTime.ParseExact(dtEmi, "yyyyMMdd", CultureInfo.InvariantCulture);
                documents.DocDueDate = DateTime.ParseExact(dtVenc, "yyyyMMdd", CultureInfo.InvariantCulture);
                documents.CardCode = codParc;
                documents.SequenceSerial = int.Parse(numNF);
                documents.SeriesString = serieNF;
                documents.SequenceModel = sequenceModel;
                documents.SequenceCode = -2;
                documents.Lines.ItemCode = itemParam;
                documents.Lines.UnitPrice = Convert.ToDouble(comissao.ToString(), new CultureInfo("pt-BR"));
                documents.Lines.Usage = util;
                if (checkRet)
                    documents.Lines.TaxLiable = SAPbobsCOM.BoYesNoEnum.tYES;
                else
                    documents.Lines.TaxLiable = SAPbobsCOM.BoYesNoEnum.tNO;
                documents.BPL_IDAssignedToInvoice = 1;
                if(documents.Add() != 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao incluir Esboço: {_api.Company.GetLastErrorDescription()}");
                    return;
                }
                string newDoc = _api.Company.GetNewObjectKey();
                pagarComissões(newDoc, comissao);
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao gravar esboço: {ex.Message}");
            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Esboço gravado...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                form.Freeze(false);
            }
        }

        private void pagarComissões(string newDoc, double comissao)
        {
            DataTable dt = form.DataSources.DataTables.Item("DT_0");
            Documents documents = (Documents)_api.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            double valorComissao = double.Parse(((EditText)form.Items.Item("Item_12").Specific).Value);
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pagando comissões...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            form.Freeze(true);
            try
            {
                for (int index = 0; index < notas.Count; index++)
                {
                    int nota = notas.ElementAt(index);
                    documents.GetByKey(nota);
                    documents.UserFields.Fields.Item("U_FOC_StatusComissao").Value = "Y";
                    documents.UserFields.Fields.Item("U_FOC_ValorComissao").Value = Convert.ToDouble(comissao.ToString(), new CultureInfo("pt-BR"));
                    documents.UserFields.Fields.Item("U_FOC_DocEntryNFCom").Value = newDoc;

                    if (documents.Update() == 0)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Comissões pagas!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Erro ao pagar comissões:{ex.Message}");
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void Listar()
        {
            form.Freeze(true);
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pesquisando dados...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                DataTable dt = form.DataSources.DataTables.Item("DT_0");
                Grid grid = (Grid)form.Items.Item("Item_9").Specific;
                string vendedor = ((EditText)this.form.Items.Item("Item_1").Specific).Value;
                string dataini = ((EditText)this.form.Items.Item("Item_5").Specific).Value;
                string datafim = ((EditText)this.form.Items.Item("Item_6").Specific).Value;
                //dataini = dataini.Substring(0, 4) + "/" + dataini.Substring(4, 2) + "/" + dataini.Substring(6, 2);
                //datafim = datafim.Substring(0, 4) + "/" + datafim.Substring(4, 2) + "/" + datafim.Substring(6, 2);
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
                string query = Queries.Comissões_Individuais.Replace("Dataini", dataini).Replace("Datafim", datafim).Replace("Vendedor", vendedor);
                dt.ExecuteQuery(query);
                grid.Columns.Item("Check").Type = BoGridColumnType.gct_CheckBox;
                for (int index = 0; index < grid.Columns.Count; index++)
                    grid.Columns.Item(index).Editable = false;
                grid.Columns.Item("Check").Editable = true;
                ((EditText)form.Items.Item("Item_12").Specific).Value = "";
                ((EditText)form.Items.Item("Item_17").Specific).Value = "";
                ((EditText)form.Items.Item("Item_18").Specific).Value = "";
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
        private void Limpar()
        {
            try
            {
                form.Freeze(true);
                DataTable dt = form.DataSources.DataTables.Item("DT_0");

                ((EditText)form.Items.Item("Item_1").Specific).Value = "";
                ((EditText)form.Items.Item("Item_2").Specific).Value = "";
                ((EditText)form.Items.Item("Item_5").Specific).Value = "";
                ((EditText)form.Items.Item("Item_6").Specific).Value = "";
                ((EditText)form.Items.Item("Item_12").Specific).Value = "";
                ((EditText)form.Items.Item("Item_16").Specific).Value = "";
                ((EditText)form.Items.Item("Item_17").Specific).Value = "";
                ((EditText)form.Items.Item("Item_18").Specific).Value = "";
                ((EditText)form.Items.Item("Item_20").Specific).Value = "";
                ((EditText)form.Items.Item("Item_21").Specific).Value = "";
                ((EditText)form.Items.Item("Item_26").Specific).Value = "";
                dt.Clear();
            }
            catch { }
            finally
            {
                form.Freeze(false);
            }

        }
        private void MarcarDesmarcarTodos()
        {
            bool marcado = !((SAPbouiCOM.CheckBox)form.Items.Item("Item_10").Specific).Checked;
            MarcarTodos(marcado);
            SomaTotal();
        }

        private void SomaTotal()
        {
            DataTable dt = form.DataSources.DataTables.Item("DT_0");
            Grid grid = (Grid)form.Items.Item("Item_9").Specific;

            double val = 0;
            double val2 = 0;
            for (int i = 0; i < grid.Rows.Count; i++)
            {
                if(grid.DataTable.Columns.Item("Check").Cells.Item(i).Value.ToString() == "Y")
                {
                    //Coluna Comissão
                    totalValue += Convert.ToDouble(dt.GetValue("Comissao", i).ToString(), new CultureInfo("pt-BR"));
                    val = Math.Round(totalValue, 2);
                    ((EditText)form.Items.Item("Item_12").Specific).Value = val.ToString(new CultureInfo("en-US"));

                    //Coluna Total
                    totalValue2 += Convert.ToDouble(dt.GetValue("Total", i).ToString(), new CultureInfo("pt-BR"));
                    val2 = Math.Round(totalValue2, 2);
                    ((EditText)form.Items.Item("Item_26").Specific).Value = val2.ToString(new CultureInfo("en-US"));

                }
                else
                {
                    totalValue -= Convert.ToDouble(grid.DataTable.Columns.Item("Comissao").Cells.Item(i).Value.ToString(), new CultureInfo("pt-BR"));
                    val = Math.Round(totalValue, 2);
                    ((EditText)form.Items.Item("Item_12").Specific).Value = val.ToString(new CultureInfo("en-US"));

                    totalValue2 -= Convert.ToDouble(grid.DataTable.Columns.Item("Total").Cells.Item(i).Value.ToString(), new CultureInfo("pt-BR"));
                    val2 = Math.Round(totalValue2, 2);
                    ((EditText)form.Items.Item("Item_26").Specific).Value = val.ToString(new CultureInfo("en-US"));
                }
            }
        }

        private void MarcarTodos(bool marcado)
        {
            form.Freeze(true);
            try
            {
                DataTable dt = form.DataSources.DataTables.Item("DT_0");
                string valor = marcado ? "Y" : "N";
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    dt.SetValue("Check", row, valor);
                }
            }
            catch (Exception)
            {

            }
            finally
            {
                form.Freeze(false);
            }
        }
    }
}
