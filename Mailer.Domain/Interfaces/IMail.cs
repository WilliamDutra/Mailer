using Mailer.Domain.Constantes;
using Mailer.Model.Entidades;
using Mailer.Model.ValueObjects;
using System;
using System.Collections.Generic;
using System.Text;

namespace Mailer.Domain.Interfaces
{
    public interface IMail
    {
        RetornoDoEnvioDoEmail Enviar(string Para, string Assunto, string Corpo);

        RetornoDoEnvioDoEmail EnviarComAnexo(string Para, string Assunto, string Corpo, string Anexo, bool Html);

        RetornoDoEnvioDoEmail EnviarComHTML(string Para, string Assunto, string Corpo);

        RetornoDoEnvioDoEmail EnviarComImagemIncorporada(string Para, string Assunto, string Anexo, string Corpo);

        List<Mensagem> LerCaixaDeEntrada(SERVICO servico = SERVICO.OUTLOOK);

    }
}
