using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADDON_PARAFLU.diapi
{
    internal sealed class API : IAPI
    {
        public Company? Company { get; set; }

        public void SetCompany(Company company)
        {
            Company = company;
        }
    }
}
