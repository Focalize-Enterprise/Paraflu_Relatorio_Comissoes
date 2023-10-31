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

namespace ADDON_PARAFLU.Uteis
{
    public sealed class Email : IEmail
    {
        internal string Name { get; set; }
        internal string E_mail { get; set; }
        internal string Senha { get; set; }
        internal string Host { get; set; }
        internal int Porta { get; set; }

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
            Name = recordset.Fields.Item("Name").Value.ToString();
            Senha = recordset.Fields.Item("U_senha").Value.ToString();
            Porta = (int)recordset.Fields.Item("U_porta").Value;
        }

        public void EnviarPorEmail(string destinationName, string destinationEmail, string[] anexos, string body, bool sendToSelf = false)
        {
            SmtpClient smtp = new SmtpClient(Host, Porta);
            try
            {
                // como a gente vai usar isso nos dois formulários é melhor colocar isso em uma outra classe 
                var mimeMessage = new MimeMessage();
                mimeMessage.From.Add(new MailboxAddress(Name, E_mail));
                mimeMessage.To.Add(new MailboxAddress(destinationName, destinationEmail));
                mimeMessage.Subject = "Comissões";
                var builder = new BodyBuilder();
                builder.HtmlBody = body;
                mimeMessage.Body = builder.ToMessageBody();
                // Convert the MimeMessage to a MailMessage
                var headers = mimeMessage.Headers;
                var mailMessage = new MailMessage();

                foreach (var recipient in mimeMessage.To)
                {
                    mailMessage.To.Add(recipient.ToString());
                }
                foreach (var header in headers)
                {
                    mailMessage.Headers.Add(header.Field, header.Value);
                }
                foreach (string anexo in anexos)
                    mailMessage.Attachments.Add(new Attachment(anexo));
                mailMessage.From = new MailAddress(mimeMessage.From[0].ToString());
                mailMessage.Subject = mimeMessage.Subject;
                mailMessage.Body = mimeMessage.HtmlBody;
                mailMessage.IsBodyHtml = true;
                smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = true;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Credentials = new NetworkCredential(E_mail, Senha);
                smtp.Send(mailMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // descarta o socket
                smtp.Dispose();
            }
        }
    }
}