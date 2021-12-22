using Mailer.Domain.Interfaces;
using Mailer.Model.Entidades;
using Mailer.Model.ValueObjects;
using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
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

        /// <summary>
        /// Método que efetuada o dispardo de e-mail
        /// </summary>
        /// <param name="Para">E-mail do destinatário</param>
        /// <param name="Assunto">Assunto do e-mail</param>
        /// <param name="Corpo">Corpo/Mensagem do e-mail</param>
        /// <returns></returns>
        public RetornoDoEnvioDoEmail Enviar(string Para, string Assunto, string Corpo)
        {
            try
            {
                return ExecutarEnvio(Para, Assunto, Corpo);
            }
            catch (Exception)
            {
                throw;

            }
        }

        /// <summary>
        /// Método que efetuada o dispardo de e-mail com um anexo
        /// </summary>
        /// <param name="Para">E-mail do destinatário</param>
        /// <param name="Assunto">Assunto do e-mail</param>
        /// <param name="Corpo">Corpo/Mensagem do e-mail</param>
        /// <param name="Arquivo">Caminho do arquivo que deseja anexar</param>
        /// <param name="Html">O corpo do e-mail contém algum HTML</param>
        /// <returns></returns>
        public RetornoDoEnvioDoEmail EnviarComAnexo(string Para, string Assunto, string Corpo, string Arquivo, bool Html = false)
        {
            try
            {
                return ExecutarEnvio(Para, Assunto, Corpo, Arquivo, Html);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Método que efetuada o dispardo de e-mail com um anexo
        /// </summary>
        /// <param name="Para">E-mail do destinatário</param>
        /// <param name="Assunto">Assunto do e-mail</param>
        /// <param name="Corpo">Corpo/Mensagem do e-mail</param>
        /// <returns></returns>
        public RetornoDoEnvioDoEmail EnviarComHTML(string Para, string Assunto, string Corpo)
        {
            try
            {
                return ExecutarEnvio(Para, Assunto, Corpo, null, true);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public List<Mensagem> LerCaixaDeEntrada()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            try
            {
                using (Pop3Client pop3Client = new Pop3Client())
                {

                    pop3Client.Connect(_Dominio, 995, true);
                    pop3Client.Authenticate(_Email, _Senha);

                    int totalCaixaEntrada = pop3Client.GetMessageCount();

                    List<Mensagem> mensagens = new List<Mensagem>();

                    for (int i = totalCaixaEntrada; i > 0; i--)
                    {
                        var msg = pop3Client.GetMessage(i);

                        mensagens.Add(new Mensagem
                        {
                            Corpo = BuscarTextoNaMensagem(msg),
                            De = msg.Headers.From.Address,
                            Para = msg.Headers.To.FirstOrDefault().Address,
                            DataEnvio = msg.Headers.DateSent,
                            Data = DateTime.Parse(msg.Headers.Date)
                        });

                    }

                    return mensagens;

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private RetornoDoEnvioDoEmail ExecutarEnvio(string Para, string Assunto, string Corpo, string Anexo = null, bool Html = false)
        {
            var retorno = new RetornoDoEnvioDoEmail();

            try
            {
                Attachment Anexado = null;
                MailMessage message = new MailMessage(_Email, Para);
                message.IsBodyHtml = Html;
                message.Subject = Assunto;
                message.Body = Corpo;

                _Smtp = new SmtpClient();
                _Smtp.EnableSsl = true;
                _Smtp.Host = _Dominio;
                _Smtp.UseDefaultCredentials = false;
                _Smtp.Credentials = new NetworkCredential(_Email, _Senha);
                //_Smtp.TargetName = "STARTTLS/smtp.office365.com";
                _Smtp.Port = _Porta;
                _Smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                if (!string.IsNullOrEmpty(Anexo))
                {
                    Anexado = new Attachment(Anexo);
                    message.Attachments.Add(Anexado);

                }

                _Smtp.Send(message);

                retorno.Sucesso = true;
                retorno.Mensagem = "E-mail enviado com sucesso!";

            }
            catch (Exception ex)
            {
                retorno.Sucesso = false;
                retorno.Mensagem = "Houve um erro ao enviar o e-mail!";
                retorno.Excecao = ex;
            }

            return retorno;
        }

        private string BuscarTextoNaMensagem(Message mensagem)
        {
            try
            {
                var Texto = mensagem.FindFirstPlainTextVersion();
                return Texto?.GetBodyAsText();
            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
