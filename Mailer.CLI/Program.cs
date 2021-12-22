using Mailer.Domain.Interfaces;
using Mailer.Domain.Services;
using Mailer.Model.ValueObjects;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Linq;

namespace Mailer.CLI
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                var serviceCollection = new ServiceCollection();
                Startup.Register(serviceCollection);

                var retorno = Run(args);

                if (!retorno.Sucesso)
                    Console.WriteLine($"{retorno.Mensagem} \n {retorno.Excecao.Message}");

            }
            catch (Exception)
            {

                throw;
            }

        }

        private static RetornoDoEnvioDoEmail Run(string[] args)
        {
            int totalArgs = args.Count();
            var mail = Startup.GetService<IMail>();

            string Para = args[0];
            string Assunto = args[1];
            string Corpo = args[2];

            return mail.Enviar(Para, Assunto, Corpo);

        }
    }
}
