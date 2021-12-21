﻿using System;
using System.Collections.Generic;
using System.Text;

namespace Mailer.Domain.Interfaces
{
    public interface IMail
    {
        bool Enviar(string Para, string Assunto, string Corpo);

        bool EnviarComAnexo(string Para, string Assunto, string Corpo, string Anexo);

        bool EnviarComHTML(string Para, string Assunto, string Corpo);

    }
}