using System;
using System.Collections.Generic;
using System.Text;

namespace Mailer.Model.ValueObjects
{
    public class RetornoDoEnvioDoEmail
    {
        public bool Sucesso { get; set; }

        public string Mensagem { get; set; }

        public Exception Excecao { get; set; }

    }
}
