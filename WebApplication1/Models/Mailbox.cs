using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class Mailbox
    {
        public string Identity { get; set; }

        public string Database { get; set; }

        public IEnumerable<string> GroupNames { get; set; }
    }
}