using System;
using System.Collections.Generic;
using System.Text;

namespace Mailer.Model.Entidades
{
    public class Mensagem
    {
        public string De { get; set; }

        public string Para { get; set; }

        public string Corpo { get; set; }

        public bool Html { get; set; }

        public DateTime DataEnvio { get; set; }

        public DateTime Data { get; set; }

    }
}
