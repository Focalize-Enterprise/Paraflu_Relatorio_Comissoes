namespace ADDON_PARAFLU.Uteis.Interfaces
{
    public interface IPDFs
    {
        string GeraPDF(string periodo1, string periodo2, string cardCode, string DBuser, string DBsenha, string caminho = "");
    }
}