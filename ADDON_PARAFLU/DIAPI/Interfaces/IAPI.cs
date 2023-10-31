using SAPbobsCOM;

namespace ADDON_PARAFLU.DIAPI.Interfaces
{
    public interface IAPI
    {
        public Company? Company { get; set; }

        void SetCompany(Company company);
    }
}