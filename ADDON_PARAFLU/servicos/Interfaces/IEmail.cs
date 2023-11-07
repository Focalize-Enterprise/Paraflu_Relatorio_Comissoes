namespace ADDON_PARAFLU.servicos.Interfaces
{
    public interface IEmail
    {
        void EnviarPorEmail(string destinationName, string destinationEmail, string[] anexos);
        void GetParamEmail();
    }
}