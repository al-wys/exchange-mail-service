using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace WebApplication1.App_Start
{
    public static class ExchangeConfig
    {
        public static string ServerFqdn { get; private set; }

        public static void Init()
        {
            ServerFqdn = ConfigurationManager.AppSettings["ExchangeServerFqdn"];
        }
    }
}