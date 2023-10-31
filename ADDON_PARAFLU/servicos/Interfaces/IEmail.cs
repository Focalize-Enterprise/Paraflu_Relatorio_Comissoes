namespace ADDON_PARAFLU.servicos.Interfaces
{
    public interface IEmail
    {
        void EnviarPorEmail(string destinationName, string destinationEmail, string[] anexos, string body, bool sendToSelf = false);
        void GetParamEmail();
    }
}