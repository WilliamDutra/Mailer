using Mailer.Domain.Interfaces;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace Mailer.Domain.Services
{
    public class MailService : IMail
    {

        private SmtpClient _Smtp;

        private string _Email;

        private string _Senha;

        private string _Dominio;

        private int _Porta;

        public MailService(string Email, string Senha, string Dominio, int Porta) 
        {
            _Email = Email;
            _Senha = Senha;
            _Dominio = Dominio;
            _Porta = Porta;
        }

        public bool Enviar(string Para, string Assunto, string Corpo)
        {
            bool retorno = false;

            try
            {
                _Smtp = new SmtpClient();               
                _Smtp.EnableSsl = true;
                _Smtp.Host = _Dominio;//"smtp.office365.com";
                _Smtp.UseDefaultCredentials = false;
                _Smtp.Credentials = new NetworkCredential(_Email, _Senha);
                _Smtp.TargetName = "STARTTLS/smtp.office365.com";
                _Smtp.Port = _Porta;//587;
                _Smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                MailMessage message = new MailMessage(_Email, Para);
                message.Subject = Assunto;
                message.Body = Corpo;

                _Smtp.Send(message);

                retorno = true;
            }
            catch (Exception)
            {

                return false;
            }

            return retorno;
        }

        public bool EnviarComAnexo(string Para, string Assunto, string Corpo, string Arquivo)
        {
            bool retorno = false;

            try
            {
                _Smtp = new SmtpClient();
                _Smtp.EnableSsl = true;
                _Smtp.Host = _Dominio;
                _Smtp.UseDefaultCredentials = false;
                _Smtp.Credentials = new System.Net.NetworkCredential(_Email, _Senha);
                _Smtp.TargetName = "STARTTLS/smtp.office365.com";
                _Smtp.Port = _Porta;
                _Smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                Attachment Anexo = new Attachment(Arquivo);

                MailMessage message = new MailMessage(_Email, Para);
                message.Subject = Assunto;
                message.Body = Corpo;
                message.Attachments.Add(Anexo);

                _Smtp.Send(message);

                
            }
            catch (Exception)
            {

                throw;
            }

            return retorno;
        }

        public bool EnviarComHTML(string Para, string Assunto, string Corpo)
        {
            bool retorno = false;

            try
            {
                _Smtp = new SmtpClient();
                _Smtp.EnableSsl = true;
                _Smtp.Host = _Dominio;//"smtp.office365.com";
                _Smtp.UseDefaultCredentials = false;
                _Smtp.Credentials = new System.Net.NetworkCredential(_Email, _Senha);
                _Smtp.TargetName = "STARTTLS/smtp.office365.com";
                _Smtp.Port = _Porta;
                _Smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                MailMessage message = new MailMessage(_Email, Para);
                message.Subject = Assunto;
                message.Body = Corpo;
                message.IsBodyHtml = true;

                _Smtp.Send(message);

                retorno = true;
            }
            catch (Exception)
            {

                retorno = false;
            }

            return retorno;

        }

    }
}
