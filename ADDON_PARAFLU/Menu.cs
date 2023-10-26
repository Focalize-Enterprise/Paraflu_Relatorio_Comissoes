using SAPbouiCOM.Framework; 
using System;
using Microsoft.Extensions.DependencyInjection;
using ADDON_PARAFLU.FORMS.UserForms;
using Application = SAPbouiCOM.Framework.Application;

namespace ADDON_PARAFLU
{
    public sealed class Menu : IMenu
    {
        //static SAPbouiCOM.Form MainMenu;
        private readonly IServiceProvider _serviceProvider;
        public Menu(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }
        /// <summary>
        /// Creates the SAPs Menus
        /// </summary>
        public void AddMenuItems()
        {
            SAPbouiCOM.MenuCreationParams oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            SAPbouiCOM.MenuItem oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "Comissoes.menu";
            oCreationPackage.String = "Relacionamento - Premiação Vendedores e Profissionais";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;
            SAPbouiCOM.Menus oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            {

            }

            try
            {
                oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("Comissoes.menu");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EnvioComissoes_Form";
                oCreationPackage.String = " Venderores - Envio de Relatórios";
                oMenus.AddEx(oCreationPackage);

            }
            catch (Exception)
            { //  Menu already exists
                //Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                //get the clicked menu to get the corect menu its necessary to go SAP and use their tool
                // and find the module you wanto to catch.
                if (pVal.BeforeAction)
                {
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        //MANUTENÇÃO DE COMISSÕES
                        case "EnvioComissoes_Form":
                            {
                                _ = _serviceProvider.GetRequiredService<EnvioDeRelatorioPorComissoes>();
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        public void RemoveMenus()
        {
            if (Application.SBO_Application.Menus.Exists("EnvioComissoes_Form"))
                Application.SBO_Application.Menus.RemoveEx("EnvioComissoes_Form");
        }
    }
}
