using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADDON_PARAFLU.Uteis
{
    internal class Vendedores
    {
        public string Code { get; set; }
        public string E_Mail { get; set; }

        public Vendedores()
        {
            Code = string.Empty;
            E_Mail = string.Empty;
        }
    }
}
