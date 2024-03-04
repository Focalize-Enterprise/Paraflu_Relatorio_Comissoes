using MimeKit;
using SAPbobsCOM;
using System;
using System.Net;
using Framework = SAPbouiCOM.Framework;
using System.Net.Mail;
using Attachment = System.Net.Mail.Attachment;
using ADDON_PARAFLU.FORMS.Recursos;
using ADDON_PARAFLU.servicos.Interfaces;
using ADDON_PARAFLU.DIAPI.Interfaces;
using ADDON_PARAFLU.Services;
using SAPbouiCOM;

namespace ADDON_PARAFLU.Uteis
{
    public sealed class Email : IEmail
    {
        internal string Name { get; set; }
        internal string E_mail { get; set; }
        internal string Senha { get; set; }
        internal string Host { get; set; }
        internal int Porta { get; set; }
        internal bool SSL { get; set; }

        private readonly IAPI _api;

        public Email(IAPI api)
        {
            _api = api;
        }

        public void GetParamEmail()
        {
            Recordset recordset = (Recordset)_api.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = Queries.Param_Email;
            recordset.DoQuery(query);
            if (recordset.RecordCount < 1)
                throw new Exception("Não há nenhum valor parametrizado para o Email");

                E_mail = recordset.Fields.Item("U_Email").Value.ToString();
                Host = recordset.Fields.Item("U_host").Value.ToString();
                Name = recordset.Fields.Item("U_Nome").Value.ToString();
                Senha = Security.Decrypt(recordset.Fields.Item("U_senha").Value.ToString());
                Porta = (int)recordset.Fields.Item("U_porta").Value;
                SSL = recordset.Fields.Item("U_SSL").Value.ToString() == "Y";
        }

        public void EnviarPorEmail(string destinationName, string destinationEmail, string[] anexos, string body, bool teste)
        {
            SmtpClient smtp = new SmtpClient(Host, Porta);
            var mimeMessage = new MimeMessage();
            var mailMessage = new MailMessage();
            try
            {
                // como a gente vai usar isso nos dois formulários é melhor colocar isso em uma outra classe
                if (string.IsNullOrEmpty(E_mail))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Nenhum Email cadastrado para envio!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }
                mimeMessage.From.Add(new MailboxAddress(Name, E_mail));
                if (teste)
                    mimeMessage.To.Add(new MailboxAddress(Name, E_mail));
                else
                    mimeMessage.To.Add(new MailboxAddress(destinationName, destinationEmail));
                mimeMessage.Subject = "Comissões";
                var builder = new BodyBuilder();
                builder.HtmlBody = body;
                mimeMessage.Body = builder.ToMessageBody();
                // Convert the MimeMessage to a MailMessage
                var headers = mimeMessage.Headers;

                foreach (var recipient in mimeMessage.To)
                {
                    mailMessage.To.Add(recipient.ToString());
                }
                foreach (var header in headers)
                {
                    mailMessage.Headers.Add(header.Field, header.Value);
                }
                foreach (string anexo in anexos.Where(x => x != ""))
                    mailMessage.Attachments.Add(new Attachment(anexo));
                mailMessage.From = new MailAddress(mimeMessage.From[0].ToString());
                mailMessage.Subject = mimeMessage.Subject;
                mailMessage.Body = mimeMessage.HtmlBody;
                mailMessage.IsBodyHtml = true;
                smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = SSL;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Credentials = new NetworkCredential(E_mail, Senha);
                smtp.Send(mailMessage);
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Erro ao enviar email:{ex.Message}");
            }
            finally
            {
                // descarta o socket
                smtp.Dispose();
                mailMessage.Attachments.Dispose();
                mimeMessage.Dispose();
                mailMessage.Dispose();
            }
        }
    }
}