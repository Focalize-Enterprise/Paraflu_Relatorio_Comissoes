using SAPbouiCOM;
using Framework = SAPbouiCOM.Framework;
using System;
using SAPbobsCOM;
using ADDON_PARAFLU.DIAPI;
using ADDON_PARAFLU.FORMS.Recursos;
using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.Services;

namespace ADDON_PARAFLU.Forms.UserForms
{
    internal class DbCredentials
    {
        SAPbouiCOM.Form form;
        private readonly IAPI _api;

        public DbCredentials(IAPI api)
        {
            _api = api;
            string xmlFormCode = Recursos.DBCredential.ToString();
            try
            {
                FormCreationParams cp = ((FormCreationParams)(Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams)));
                cp.FormType = "DbCredentials";
                cp.XmlData = xmlFormCode;
                cp.UniqueID = "DbCredentials";
                form = Framework.Application.SBO_Application.Forms.AddEx(cp);
                CustomInitialize();
            }
            catch (Exception ex)
            {
                Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao crair o formulário: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            Recordset recordset = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recordset1 = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = @"SELECT ""U_User"", ""U_Pass"", ""U_Past"", ""U_Item"", ""U_Util"", ""U_Crystal"" FROM ""@FOC_DB_CONF"" WHERE ""Code"" = '1'";
            string query1 = @"SELECT ""U_Email"", ""U_senha"", ""U_host"", ""U_porta"", ""U_Nome"", ""U_SSL"" FROM ""@FOC_EMAIL_PARAM"" WHERE ""Code"" = '1'";
            recordset.DoQuery(query);
            recordset1.DoQuery(query1);
            if (recordset.RecordCount > 0)
            {
                ((EditText)form.Items.Item("Item_0").Specific).Value = recordset.Fields.Item(0).Value.ToString();
                ((EditText)form.Items.Item("Item_1").Specific).Value = Security.Decrypt(recordset.Fields.Item(1).Value.ToString());
                ((EditText)form.Items.Item("Item_3").Specific).Value = recordset.Fields.Item(2).Value.ToString();
                ((EditText)form.Items.Item("Item_23").Specific).Value = recordset.Fields.Item(3).Value.ToString();
                ((EditText)form.Items.Item("Item_24").Specific).Value = recordset.Fields.Item(4).Value.ToString();
                ((EditText)form.Items.Item("Item_26").Specific).Value = recordset.Fields.Item(5).Value.ToString();

            }
            if (recordset1.RecordCount > 0)
            {
                ((EditText)form.Items.Item("Item_11").Specific).Value = recordset1.Fields.Item(0).Value.ToString();
                ((EditText)form.Items.Item("Item_12").Specific).Value = Security.Decrypt(recordset1.Fields.Item(1).Value.ToString());
                ((EditText)form.Items.Item("Item_17").Specific).Value = recordset1.Fields.Item(3).Value.ToString();
                ((EditText)form.Items.Item("Item_18").Specific).Value = recordset1.Fields.Item(2).Value.ToString();
                ((EditText)form.Items.Item("Item_27").Specific).Value = recordset1.Fields.Item(4).Value.ToString();
                ((SAPbouiCOM.CheckBox)form.Items.Item("Item_29").Specific).Checked = recordset1.Fields.Item(5).Value.ToString() == "Y";
            }

            Framework.Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (FormUID != form.UniqueID)
                return;
            try
            {

                if(pVal.EventType == BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                    Framework.Application.SBO_Application.ItemEvent -= SBO_Application_ItemEvent;
                if(pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction && pVal.ItemUID == "1")
                {
                    bool checkSSL = ((SAPbouiCOM.CheckBox)form.Items.Item("Item_29").Specific).Checked;
                    UserTable userTable = _api.Company.UserTables.Item("FOC_DB_CONF");
                    UserTable userTable1 = _api.Company.UserTables.Item("FOC_EMAIL_PARAM");
                    if (userTable.GetByKey("1"))
                    {
                        userTable.UserFields.Fields.Item("U_User").Value = ((EditText)form.Items.Item("Item_0").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Pass").Value = ((EditText)form.Items.Item("Item_1").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Past").Value = ((EditText)form.Items.Item("Item_3").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Item").Value = ((EditText)form.Items.Item("Item_23").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Crystal").Value = ((EditText)form.Items.Item("Item_26").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Util").Value = ((EditText)form.Items.Item("Item_24").Specific).Value;
                        if (userTable.Update() != 0)
                        {
                            Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao tentar atualizar os dados: {_api.Company.GetLastErrorDescription()}");
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        userTable.Code = "1";
                        userTable.Name = "Conf";
                        userTable.UserFields.Fields.Item("U_User").Value = ((EditText)form.Items.Item("Item_0").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Pass").Value = ((EditText)form.Items.Item("Item_1").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Past").Value = ((EditText)form.Items.Item("Item_3").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Item").Value = ((EditText)form.Items.Item("Item_23").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Crystal").Value = ((EditText)form.Items.Item("Item_26").Specific).Value;
                        userTable.UserFields.Fields.Item("U_Util").Value = ((EditText)form.Items.Item("Item_24").Specific).Value;

                        if (userTable.Add() != 0)
                        {
                            Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao tentar adicionar os dados: {_api.Company.GetLastErrorDescription()}");
                            BubbleEvent = false;
                        }
                    }

                    if (userTable1.GetByKey("1"))
                    {
                        userTable1.UserFields.Fields.Item("U_Nome").Value = ((EditText)form.Items.Item("Item_27").Specific).Value;
                        userTable1.UserFields.Fields.Item("U_Email").Value = ((EditText)form.Items.Item("Item_11").Specific).Value;
                        userTable1.UserFields.Fields.Item("U_senha").Value = Security.Encrypt(((EditText)form.Items.Item("Item_12").Specific).Value);
                        userTable1.UserFields.Fields.Item("U_host").Value = ((EditText)form.Items.Item("Item_18").Specific).Value;
                        userTable1.UserFields.Fields.Item("U_porta").Value = ((EditText)form.Items.Item("Item_17").Specific).Value;

                        if (checkSSL)
                            userTable1.UserFields.Fields.Item("U_SSL").Value = "Y";
                        else
                            userTable1.UserFields.Fields.Item("U_SSL").Value = "N";

                        if (userTable1.Update() != 0)
                        {
                            Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao tentar atualizar os dados: {_api.Company.GetLastErrorDescription()}");
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        userTable1.Code = "1";
                        userTable1.Name = "Conf";
                        userTable1.UserFields.Fields.Item("U_Nome").Value = ((EditText)form.Items.Item("Item_27").Specific).Value;
                        userTable1.UserFields.Fields.Item("U_Email").Value = ((EditText)form.Items.Item("Item_11").Specific).Value;
                        userTable1.UserFields.Fields.Item("U_senha").Value = Security.Encrypt(((EditText)form.Items.Item("Item_12").Specific).Value);
                        userTable1.UserFields.Fields.Item("U_host").Value = ((EditText)form.Items.Item("Item_18").Specific).Value;
                        userTable1.UserFields.Fields.Item("U_porta").Value = ((EditText)form.Items.Item("Item_17").Specific).Value;

                        if (checkSSL)
                            userTable1.UserFields.Fields.Item("U_SSL").Value = "Y";
                        else
                            userTable1.UserFields.Fields.Item("U_SSL").Value = "N";

                        if (userTable1.Add() != 0)
                        {
                            Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao tentar adicionar os dados: {_api.Company.GetLastErrorDescription()}");
                            BubbleEvent = false;
                        }
                    }
                }
            
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"erro {ex.Message}");
            }
        }
    }
}
