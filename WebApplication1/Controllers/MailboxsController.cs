using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.RegularExpressions;
using System.Web.Http;
using WebApplication1.App_Start;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class MailboxsController : ApiController
    {
        private WSManConnectionInfo exchangeConnectionInfo;
        private WSManConnectionInfo ExchangeConnectionInfo
        {
            get
            {
                if (exchangeConnectionInfo == null)
                {
                    exchangeConnectionInfo = new WSManConnectionInfo(new Uri($"http://{ExchangeConfig.ServerFqdn}/PowerShell/"),
                        "http://schemas.microsoft.com/powershell/Microsoft.Exchange", PSCredential.Empty);
                }

                return exchangeConnectionInfo;
            }
        }

        // POST api/mailboxs
        public string Post([FromBody] Mailbox mailbox)
        {
            // Connect to PowerShell of Exchange Server
            using (Runspace runspace = RunspaceFactory.CreateRunspace(ExchangeConnectionInfo))
            {
                runspace.Open();

                string response = string.Empty;

                // Enable mailbox
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = runspace;

                    ps.AddCommand("Enable-Mailbox").AddParameter("Identity", mailbox.Identity);
                    if (!string.IsNullOrEmpty(mailbox.Database))
                    {
                        ps.AddParameter("Database", mailbox.Database);
                    }
                    try
                    {
                        System.Collections.ObjectModel.Collection<PSObject> enableResult = ps.Invoke();

                        foreach (var item in enableResult)
                        {
                            response += "Name: " + item.Properties["Name"].Value;
                            response += "; ServerName: " + item.Properties["ServerName"].Value;
                        }
                    }
                    catch (Exception)
                    {
                        // Error when enable mailbox
                        throw;
                    }
                }

                // Add to groups
                if (mailbox.GroupNames != null)
                {
                    List<string> groupNames = new List<string>(mailbox.GroupNames);
                    if (groupNames.Count > 0)
                    {
                        using (PowerShell ps = PowerShell.Create())
                        {
                            ps.Runspace = runspace;

                            for (int i = 0; i < groupNames.Count - 1; i++)
                            {
                                ps.AddCommand("Add-DistributionGroupMember").
                                    AddParameter("Identity", groupNames[i]).
                                    AddParameter("Member", mailbox.Identity).
                                    AddStatement();
                            }
                            // No statement for the last command
                            ps.AddCommand("Add-DistributionGroupMember").
                                AddParameter("Identity", groupNames[groupNames.Count - 1]).
                                AddParameter("Member", mailbox.Identity);

                            try
                            {
                                System.Collections.ObjectModel.Collection<PSObject> add2GroupResult = ps.Invoke();
                            }
                            catch (Exception)
                            {
                                // Error when adding to groups
                                throw;
                            }
                        }
                    }
                }

                return response;
            }
        }

        // DELETE api/mailboxs
        public string Delete([FromBody] Mailbox mailbox)
        {
            // Connect to PowerShell of Exchange Server
            using (Runspace runspace = RunspaceFactory.CreateRunspace(ExchangeConnectionInfo))
            {
                runspace.Open();

                // Enable mailbox
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = runspace;

                    ps.AddCommand("Disable-Mailbox").AddParameter("Identity", mailbox.Identity).AddParameter("Confirm", false);

                    try
                    {
                        System.Collections.ObjectModel.Collection<PSObject> disableResult = ps.Invoke();

                        return "Done";
                    }
                    catch (Exception)
                    {
                        // Error when enable mailbox
                        throw;
                    }
                }
            }
        }
    }
}
