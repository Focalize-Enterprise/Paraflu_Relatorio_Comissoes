using ADDON_PARAFLU.DIAPI.Interfaces;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADDON_PARAFLU.FORMS.UserForms
{
    public sealed class Form2
    {
        private readonly IAPI _api;
        private readonly string formid;
        private readonly SAPbouiCOM.Form form;
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

                            }
                        }
                        break;
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        {
                            string value = string.Empty;
                            IChooseFromListEvent CFLEvento = (IChooseFromListEvent)pVal;
                            DataTable dataTable = CFLEvento.SelectedObjects;

                            if (dataTable != null)
                                value = dataTable.GetValue(0, 0).ToString();
                            if (string.IsNullOrEmpty(value))
                                return;

                            switch (CFLEvento.ChooseFromListUID)
                            {
                                case "CFL_0":
                                    {
                                        ((EditText)this.form.Items.Item("Item_1").Specific).Value = value;
                                        ((EditText)this.form.Items.Item("Item_2").Specific).Value = dataTable.GetValue("SlpName", 0).ToString();
                                    }
                                    break;

                            }
                        }
                        break;
                }
            }
        }
    }
}
