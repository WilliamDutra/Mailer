using Mailer.Domain.Interfaces;
using Mailer.Domain.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.IO;

namespace Mailer.CLI
{
    public static class Startup
    {
        public static IConfiguration Configuration;

        private static ServiceProvider _Container;

        public static void Register(IServiceCollection serviceCollection)
        {
            var _serviceCollection = serviceCollection;


            Configuration = new ConfigurationBuilder()
                                    .SetBasePath(Directory.GetParent(AppContext.BaseDirectory).FullName)
                                    .AddJsonFile("appsettings.json", false)
                                    .Build();

            string Dominio = Configuration.GetSection("Credenciais:Dominio").Value;
            string Usuario = Configuration.GetSection("Credenciais:Usuario").Value;
            string Senha = Configuration.GetSection("Credenciais:Senha").Value;
            int Porta = Convert.ToInt32( Configuration.GetSection("Credenciais:Porta").Value);

            _serviceCollection.AddScoped<IMail, MailService>((Mail) => new MailService(Usuario, Senha, Dominio, Porta));

            _Container = _serviceCollection.BuildServiceProvider();
            
        }

        public static T GetService<T>()
        {
            return _Container.GetService<T>();
        }
        
    }
}
