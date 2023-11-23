﻿//------------------------------------------------------------------------------
// <auto-generated>
//     O código foi gerado por uma ferramenta.
//     Versão de Tempo de Execução:4.0.30319.42000
//
//     As alterações ao arquivo poderão causar comportamento incorreto e serão perdidas se
//     o código for gerado novamente.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ADDON_PARAFLU.FORMS.Recursos {
    using System;
    
    
    /// <summary>
    ///   Uma classe de recurso de tipo de alta segurança, para pesquisar cadeias de caracteres localizadas etc.
    /// </summary>
    // Essa classe foi gerada automaticamente pela classe StronglyTypedResourceBuilder
    // através de uma ferramenta como ResGen ou Visual Studio.
    // Para adicionar ou remover um associado, edite o arquivo .ResX e execute ResGen novamente
    // com a opção /str, ou recrie o projeto do VS.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Queries {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Queries() {
        }
        
        /// <summary>
        ///   Retorna a instância de ResourceManager armazenada em cache usada por essa classe.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("ADDON_PARAFLU.FORMS.Recursos.Queries", typeof(Queries).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Substitui a propriedade CurrentUICulture do thread atual para todas as
        ///   pesquisas de recursos que usam essa classe de recurso de tipo de alta segurança.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Consulta uma cadeia de caracteres localizada semelhante a Select  
        ///&apos;N&apos; as &quot;Selecionado&quot;,
        ///&quot;SlpCode&quot; as &quot;Código&quot;,
        ///&quot;SlpName&quot; as &quot;Nome do Vendedor&quot;,
        ///&quot;Email&quot; as &quot;Email do vendedor&quot;,
        ///&quot;DocDate&quot; as &quot;Data de geração&quot;,
        ///sum(&quot;LineTotal&quot;) as &quot;Total&quot;,
        ///sum(&quot;IPI&quot;) as &quot;IPI&quot;,
        ///sum(&quot;ICMSST&quot;) as &quot;ICMS-ST&quot;,
        ///sum((&quot;LineTotal&quot; + ifnull(&quot;IPI&quot;, 0) + ifnull(&quot;ICMSST&quot;, 0))*(&quot;Com&quot;) )/100 as &quot;Comissão&quot;
        /// from 
        /// 
        /// (
        ///
        ///SELECT  top 10 
        ///&apos;NS&apos; &quot;Doc&quot;,
        ///T0.&quot;ObjType&quot;,
        ///T0.&quot;DocEntry&quot;, 
        ///T0.&quot;DocDate&quot;, 
        ///T0.&quot;DocDueDate&quot;, 
        ///T0.&quot;CardCode&quot;, 
        ///T0.&quot;CardName&quot;, 
        ///T0.&quot;SlpCode&quot;, 
        ///T2.&quot;SlpName&quot;,
        ///T2.&quot;Email&quot;,
        ///T0.&quot;Serial&quot;, 
        ///T1.&quot;I [o restante da cadeia de caracteres foi truncado]&quot;;.
        /// </summary>
        internal static string Notas_Fiscais_HANA {
            get {
                return ResourceManager.GetString("Notas_Fiscais_HANA", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Consulta uma cadeia de caracteres localizada semelhante a Select  
        ///&apos;N&apos; as &quot;Selecionado&quot;,
        ///&quot;SlpCode&quot; as &quot;Código&quot;,
        ///&quot;SlpName&quot; as &quot;Nome do Vendedor&quot;,
        ///&quot;Email&quot; as &quot;Email do vendedor&quot;,
        ///&quot;DocDate&quot; as &quot;Data de geração&quot;,
        ///sum(&quot;LineTotal&quot;) as &quot;Total&quot;,
        ///sum(&quot;IPI&quot;) as &quot;IPI&quot;,
        ///sum(&quot;ICMSST&quot;) as &quot;ICMS-ST&quot;,
        ///sum((&quot;LineTotal&quot; + isnull(&quot;IPI&quot;, 0) + isnull(&quot;ICMSST&quot;, 0))*(&quot;Com&quot;) )/100 as &quot;Comissão&quot;
        /// from 
        /// 
        /// (
        ///
        ///SELECT  top 10
        ///&apos;NS&apos; &quot;Doc&quot;,
        ///T0.&quot;ObjType&quot;,
        ///T0.&quot;DocEntry&quot;, 
        ///T0.&quot;DocDate&quot;, 
        ///T0.&quot;DocDueDate&quot;, 
        ///T0.&quot;CardCode&quot;, 
        ///T0.&quot;CardName&quot;, 
        ///T0.&quot;SlpCode&quot;, 
        ///T2.&quot;SlpName&quot;,
        ///T2.&quot;Emai [o restante da cadeia de caracteres foi truncado]&quot;;.
        /// </summary>
        internal static string Notas_Fiscais_SQL {
            get {
                return ResourceManager.GetString("Notas_Fiscais_SQL", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Consulta uma cadeia de caracteres localizada semelhante a SELECT TOP 1 * FROM &quot;@FOC_EMAIL_PARAM&quot;.
        /// </summary>
        internal static string Param_Email {
            get {
                return ResourceManager.GetString("Param_Email", resourceCulture);
            }
        }
    }
}
