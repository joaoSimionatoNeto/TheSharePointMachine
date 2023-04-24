using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;


namespace The_SharePoint_Machine
{
    internal class Connection
    {
        public ClientContext login(string url, string username, string password)
        {
            try
            {
                ClientContext context = new ClientContext(url);
                SecureString SecurePassword = new SecureString();
                foreach (char c in password.ToCharArray())
                {
                    SecurePassword.AppendChar(c);
                }
                context.Credentials = new SharePointOnlineCredentials(username, SecurePassword);
                Web web = context.Web;
                context.ExecuteQuery();
                return context;
            } catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
                return null;
            }
        }
    }
}
