using SAPbobsCOM;

namespace ADDON_PARAFLU.diapi
{
    public interface IAPI
    {
        public Company? Company { get; set; }

        void SetCompany(Company company);
    }
}