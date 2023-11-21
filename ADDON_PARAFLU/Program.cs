using ADDON_PARAFLU.DIAPI;
using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.Forms.UserForms;
using ADDON_PARAFLU.FORMS.UserForms;
using ADDON_PARAFLU.servicos.Interfaces;
using ADDON_PARAFLU.Uteis;
using ADDON_PARAFLU.Uteis.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Application = SAPbouiCOM.Framework.Application;

namespace ADDON_PARAFLU
{
    public static class Program
    {
        private static IServiceProvider? ServiceProvider;

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                /// primeiro teste com dependecy inject usando o SAP B1, sujeito a varias mudanças - joão
                /// 10/08/2023 todos os formulários serão injetados como serviços
                IHost host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                 // servicos do SAP
                 // conexão com api
                    services.AddSingleton(typeof(IAPI), typeof(API));
                    services.AddSingleton(typeof(IMenu), typeof(Menu));
                    services.AddSingleton(typeof(IPDFs), typeof(PDFs));
                    // serviços usados pelos formulários
                    services.AddTransient<IEmail, Email>();
                    services.AddTransient<IPDFs, PDFs>();
                    // formularios abaixo
                    services.AddTransient<EnvioDeRelatorioPorComissoes>();
                    services.AddTransient<DbCredentials>();
                })
                .Build();

                ServiceProvider = host.Services;

                Application oApp;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }

                // valida o banco de dados
                var comp = ServiceProvider.GetRequiredService<IAPI>();
                comp.SetCompany((SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany());
                if (comp.Company is null)
                {
                    System.Windows.Forms.MessageBox.Show("Erro ao tentar conectar com a DIAPI Desligando o add-on");
                    return;
                }


                IMenu menu = ServiceProvider.GetRequiredService<IMenu>();
                menu.RemoveMenus();
                menu.AddMenuItems();
                oApp.RegisterMenuEventHandler(menu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    {
                        //Exit Add-On
                        IMenu menu = ServiceProvider!.GetRequiredService<IMenu>();
                        menu.RemoveMenus();
                        System.Windows.Forms.Application.Exit();
                    }
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}