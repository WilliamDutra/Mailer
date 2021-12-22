using Mailer.Domain.Interfaces;
using Mailer.Domain.Services;
using Microsoft.Extensions.DependencyInjection;
using System;

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
                                

            }
            catch (Exception)
            {

                throw;
            }

        }
    }
}
