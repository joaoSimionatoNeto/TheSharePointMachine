using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Reflection;
using System.Xml;
using System.IO;
using Microsoft.SharePoint.Client.Discovery;
using static System.Net.Mime.MediaTypeNames;

namespace The_SharePoint_Machine
{
    internal class Program
    {
        static void Main(string[] args)
        {
            int program = -1;
            string url = " ";
            string username;
            string password;
            string Title;
            string NewTitle;
            string Description;
            string InternalName;
            string xmlAdd;
            SP_Operations operations = null;
            Connection conn = new Connection();
            ClientContext context = null;
            string permission;
            string RemoveRole;
            string AddRole;
            string users;
            string Addusers;
            string RemoveUsers;
            string Fields;
            string file;
            string XmlPath;
            char[] delimiterChars = { ' ', ',' };




            while (program != 10)
            {
                switch (program)
                {
                    case -1:
                        while (context == null)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("*****************************************************");
                            Console.WriteLine("*              The SharePoint Machine               *");
                            Console.WriteLine("*****************************************************");
                            Console.WriteLine();
                            Console.WriteLine("Para começar insira a url do Site SharePoint em que deseja trabalhar, caso vá criar o site, insira a ulr do tenant ex: https://dominio.sharepoint.com/_layouts/15/sharepoint.aspx, mas atenção só é possivel criar o site se você for administrador do siteCollection!");
                            url = Console.ReadLine();
                            Console.WriteLine("Insira o seu usuário, é importante que ele tenha permissão de administrador");
                            username = Console.ReadLine();
                            Console.WriteLine("Insira a sua senha");
                            password = Console.ReadLine();
                            context = conn.login(url, username, password);
                        }
                        program = 0;
                        operations = new SP_Operations(context);
                        Console.Clear();
                        continue;
                    case 0:
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("*****************************************************");
                        Console.WriteLine("*              The SharePoint Machine               *");
                        Console.WriteLine("*****************************************************");
                        Console.WriteLine();
                        Console.WriteLine("Digite o número da ação que você quer executar: ");
                        Console.WriteLine();
                        Console.WriteLine("**************************************");
                        Console.WriteLine("1 - Criar site");
                        Console.WriteLine("2 - Deletar Site");
                        Console.WriteLine("3 - Permissões de Site");
                        Console.WriteLine("4 - Criar Lista ");
                        Console.WriteLine("5 - Adicionar ou remover campos da lista ");
                        Console.WriteLine("6 - Deletar Lista");
                        Console.WriteLine("7 - Adicionar Itens a Lista");
                        Console.WriteLine("8 - Remover Itens da Lista");
                        Console.WriteLine("9 - Exportar itens da Lista");
                        Console.WriteLine("10 - Sair");
                        program = Int32.Parse(Console.ReadLine());
                        continue;
                    case 1:
                        Console.WriteLine("Insira o Nome Interno do seu site");
                        InternalName = Console.ReadLine();
                        Console.WriteLine("Insira o Nome do seu site");
                        Title = Console.ReadLine();
                        Console.WriteLine("Insira a descrição do seu site");
                        Description = Console.ReadLine();
                        operations.CreateSite(InternalName, Title, Description);
                        program = -1;
                        break;
                    case 2:
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Tem certeza que deseja apagar o site: " + url + "(S | N)");
                        if(Console.ReadLine().ToUpper() == "S")
                        {
                            operations.DeleteSite();
                            Console.Clear();
                        }
                        else
                        {
                            program = 0;
                            Console.Clear();
                        }
                        break;
                    case 3:
                        program = 0;
                        while (program != 7)
                        {
                            switch (program)
                            {
                                case 0:
                                    Console.Clear();
                                    Console.WriteLine("*****************************************************");
                                    Console.WriteLine("*              The SharePoint Machine               *");
                                    Console.WriteLine("*****************************************************");
                                    Console.WriteLine();
                                    Console.WriteLine("Digite o número da ação que você quer executar: ");
                                    Console.WriteLine("1 - Criar Nível de Permissão");
                                    Console.WriteLine("2 - Editar Nível de Permissão");
                                    Console.WriteLine("3 - Remover Nível de Permissão");
                                    Console.WriteLine("4 - Criar Grupo ");
                                    Console.WriteLine("5 - Editar Grupo ");
                                    Console.WriteLine("6 - Deletar Grupo ");
                                    Console.WriteLine("7 - Voltar");
                                    program = Int32.Parse(Console.ReadLine());
                                    break;
                                case 1:
                                    Console.Clear();
                                    Console.WriteLine("*****************************************************");
                                    Console.WriteLine("*              The SharePoint Machine               *");
                                    Console.WriteLine("*****************************************************");
                                    Console.WriteLine();
                                    Console.WriteLine("Criar Níveis de Permissão");
                                    Console.WriteLine();
                                    Console.WriteLine("Insira o nome do nível de permissão: ");
                                    Title = Console.ReadLine();
                                    Console.WriteLine("Insira uma breve descrição do nível de permissão: ");
                                    Description= Console.ReadLine();
                                    Console.WriteLine();
                                    Console.WriteLine("Digite os números separados por vírgula ou apenas espaços das permissões que você deseja dar a este nível: ");
                                    Console.WriteLine("1 - Ver Itens de Listas");
                                    Console.WriteLine("2 - Adicionar Itens as Listas");
                                    Console.WriteLine("3 - Editar Itens das Listas");
                                    Console.WriteLine("4 - Deletar Itens das Listas");
                                    Console.WriteLine("5 - Aprovar Itens");
                                    Console.WriteLine("6 - Abrir Itens");
                                    Console.WriteLine("7 - Ver Versões");
                                    Console.WriteLine("8 - Deletar Versões");
                                    Console.WriteLine("9 - Cancelar Checkout");
                                    Console.WriteLine("10 - Gerenciar Views pessoais");
                                    Console.WriteLine("11 - Gerenciar Listas");
                                    Console.WriteLine("12 - Exibir Páginas de Aplicativo ");
                                    Console.WriteLine("13 - Criar Alertas ");
                                    Console.WriteLine("14 - Gerenciar Alertas ");
                                    Console.WriteLine("15 - Ver páginas ");
                                    Console.WriteLine("16 - Adicionar e customizar páginas ");
                                    Console.WriteLine("17 - Adicionar temas");
                                    Console.WriteLine("18 - Adicionar folhas de estilos");
                                    Console.WriteLine("19 - Exibir Dados do Web Analytics");
                                    Console.WriteLine("20 - Criar SubSites");
                                    Console.WriteLine("21 - Criar Grupos");
                                    Console.WriteLine("22 - Gerenciar Permissões");
                                    Console.WriteLine("23 - Pesquisar Diretórios");
                                    Console.WriteLine("24 - Editar Informações Pessoais do Usuário");
                                    Console.WriteLine("25 - Usar Recursos de Integração de Cliente");
                                    Console.WriteLine("26 - Usar Interfaces Remotas");
                                    Console.WriteLine("27 - Enumerar Permissões ");
                                    Console.WriteLine("28 - Todas as Permissões");
                                    permission = Console.ReadLine();
                                    string[] permissions = permission.Split(delimiterChars);
                                    operations.CreateRole(Title, Description, permissions);
                                    program = 0;
                                    break;
                                case 2:
                                    Console.Clear();
                                    Console.WriteLine("*****************************************************");
                                    Console.WriteLine("*              The SharePoint Machine               *");
                                    Console.WriteLine("*****************************************************");
                                    Console.WriteLine();
                                    Console.WriteLine("Editar níveis de permissão");
                                    Console.WriteLine();
                                    Console.WriteLine("Insira o nome do nível de permissão a ser editado: ");
                                    Title = Console.ReadLine();
                                    Console.WriteLine("Insira uma breve descrição do nível de permissão (caso não queira editar a descrição apenas precione enter): ");
                                    Description = Console.ReadLine();
                                    Console.WriteLine();
                                    Console.WriteLine("Digite os números separados por vírgula ou apenas espaços das permissões que você deseja adicionar a este nível: ");
                                    Console.WriteLine("1 - Ver Itens de Listas");
                                    Console.WriteLine("2 - Adicionar Itens as Listas");
                                    Console.WriteLine("3 - Editar Itens das Listas");
                                    Console.WriteLine("4 - Deletar Itens das Listas");
                                    Console.WriteLine("5 - Aprovar Itens");
                                    Console.WriteLine("6 - Abrir Itens");
                                    Console.WriteLine("7 - Ver Versões");
                                    Console.WriteLine("8 - Deletar Versões");
                                    Console.WriteLine("9 - Cancelar Checkout");
                                    Console.WriteLine("10 - Gerenciar Views pessoais");
                                    Console.WriteLine("11 - Gerenciar Listas");
                                    Console.WriteLine("12 - Exibir Páginas de Aplicativo ");
                                    Console.WriteLine("13 - Criar Alertas ");
                                    Console.WriteLine("14 - Gerenciar Alertas ");
                                    Console.WriteLine("15 - Ver páginas ");
                                    Console.WriteLine("16 - Adicionar e customizar páginas ");
                                    Console.WriteLine("17 - Adicionar temas");
                                    Console.WriteLine("18 - Adicionar folhas de estilos");
                                    Console.WriteLine("19 - Exibir Dados do Web Analytics");
                                    Console.WriteLine("20 - Criar SubSites");
                                    Console.WriteLine("21 - Criar Grupos");
                                    Console.WriteLine("22 - Gerenciar Permissões");
                                    Console.WriteLine("23 - Pesquisar Diretórios");
                                    Console.WriteLine("24 - Editar Informações Pessoais do Usuário");
                                    Console.WriteLine("25 - Usar Recursos de Integração de Cliente");
                                    Console.WriteLine("26 - Usar Interfaces Remotas");
                                    Console.WriteLine("27 - Enumerar Permissões ");
                                    Console.WriteLine("28 - Todas as Permissões");
                                    permission = Console.ReadLine();
                                    string[] AddPermissions = permission.Split(delimiterChars);
                                    Console.WriteLine();
                                    Console.WriteLine("Utilizando a mesma lista acima, insira os números das permissões que deseja remover: ");
                                    permission = Console.ReadLine();
                                    string[] RemovePermissions = permission.Split(delimiterChars);
                                    operations.EditRole(Title, Description, AddPermissions, RemovePermissions);
                                    program = 0;
                                    break;
                                case 3:
                                    Console.WriteLine("Insira o nome do nível de permissão a ser deletado:");
                                    Title = Console.ReadLine();
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Tem certeza que deseja apagar o nível de permissão: " + Title + "(S | N)");
                                    if (Console.ReadLine().ToUpper() == "S")
                                    {
                                        operations.DeleteRole(Title);
                                        Console.Clear();
                                    }
                                    else
                                    {
                                        program = 0;
                                        Console.Clear();
                                    }
                                    program = 0;
                                    break;
                                case 4:
                                    Console.WriteLine("Nome do grupo: ");
                                    Title = Console.ReadLine();
                                    Console.WriteLine("Descrição do grupo");
                                    Description = Console.ReadLine();
                                    Console.WriteLine("Digite os usuários a serem adicionados a este grupo: ");
                                    users = Console.ReadLine();
                                    users = users.Replace(",", "");
                                    string[] user = users.Split(delimiterChars);
                                    Console.WriteLine("Digite o nome do(s) grupo() de permissão a ser consedido a esse grupo: ");
                                    permission = Console.ReadLine();
                                    permission = permission.Replace(",", "");
                                    string[] Addpermissions = permission.Split(delimiterChars);
                                    operations.CreateGroup(Title, Description, user, Addpermissions);
                                    program = 0;
                                    break;
                                case 5:
                                    Console.WriteLine("Nome do grupo a ser editado:");
                                    Title = Console.ReadLine();
                                    Console.WriteLine("Digite o novo nome do grupo (caso não queira editar o nome apenas precione enter): ");
                                    NewTitle = Console.ReadLine();
                                    Console.WriteLine("Insira uma breve descrição do grupo (caso não queira editar a descrição apenas precione enter): ");
                                    Description = Console.ReadLine();
                                    Console.WriteLine("Digite o nome dos usuarios a serem adicionados separados por vírgula sem espaço ou apenas espaço (caso não queira adicionar ninguém precione enter): ");
                                    Addusers = Console.ReadLine();
                                    Addusers = Addusers.Replace(",", "");
                                    string[] AddedUsers = Addusers.Split(delimiterChars);
                                    Console.WriteLine("Digite o nome dos usuarios a serem removidos (caso não queira remover ninguém precione enter): ");
                                    RemoveUsers = Console.ReadLine();
                                    RemoveUsers = RemoveUsers.Replace(",", "");
                                    string[] RemovedUsers = RemoveUsers.Split(delimiterChars);
                                    Console.WriteLine("Digite o nome do grupo de permissão a ser removido (caso não queira remover nenhum precione enter):");
                                    RemoveRole = Console.ReadLine();
                                    RemoveRole = RemoveRole.Replace(",", "");
                                    string[] RemovedRoles = RemoveRole.Split(delimiterChars);
                                    Console.WriteLine("Digite o nome do grupo de permissão a ser adicionado (caso não queira remover nenhum precione enter):");
                                    AddRole = Console.ReadLine();
                                    string[] AddedRole = AddRole.Split(delimiterChars);
                                    operations.EditGroup(Title, NewTitle, Description, AddedUsers, RemovedUsers, RemovedRoles, AddedRole);
                                    program = 0;
                                    break;
                                case 6:
                                    Console.WriteLine("Digite o nome o grupo a ser removido");
                                    Title = Console.ReadLine();
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Tem certeza que deseja apagar o grupo de permissão: " + Title + "(S | N)");
                                    if (Console.ReadLine().ToUpper() == "S")
                                    {
                                        operations.DeleteGroup(Title);
                                        Console.Clear();
                                    }
                                    else
                                    {
                                        program = 0;
                                        Console.Clear();
                                    }
                                    break;
                                default:
                                    Console.Clear();
                                    program = 0;
                                    break;

                            }
                        }
                        program = 0;
                        Console.Clear();
                        break;
                    case 4:
                        Console.Write("Nome da lista: ");
                        Title = Console.ReadLine();
                        Console.Write("Internal name da lista: ");
                        InternalName = Console.ReadLine();
                        Console.Write("Insira a descrição da lista: ");
                        Description = Console.ReadLine();
                        Console.Write("Insira o caminho do arquivo Xml contendo os campos a serem criados na lista: ");
                        XmlPath =  Console.ReadLine();
                        operations.CreateList(Title, InternalName, Description, XmlPath);
                        break;
                    case 5:
                        Console.Write("Digite o InternalName da lista a ser editada: ");
                        InternalName = Console.ReadLine();
                        Console.Write("Digite o caminho do arquivo Xml contento os campos a serem adicionados (Caso não queira adicionar campos, precione enter): ");
                        XmlPath = Console.ReadLine();
                        Console.WriteLine("Digite os internal names separados por vírgulas, sem espaços, dos campos a serem removidos (Caso não queira adicionar campos, precione enter): ");
                        Fields = Console.ReadLine();
                        string[] field = Fields.Split(delimiterChars);
                        operations.EditList(InternalName, XmlPath, field);
                        program = 0;
                        break;
                    case 6:
                        Console.WriteLine("Digite o internal name da lista a ser deletada: ");
                        InternalName = Console.ReadLine();
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Tem certeza que deseja apagar a lista: " + InternalName + "(S | N)");
                        if (Console.ReadLine().ToUpper() == "S")
                        {
                            operations.DeleteList(InternalName);
                            Console.Clear();
                        }
                        else
                        {
                            program = 0;
                            Console.Clear();
                        }
                        break;
                    case 7:
                        Console.WriteLine("Função em desenvolvimento, retornando ao menu........");
                        program = 0;
                        break;
                    case 8:
                        Console.WriteLine("Função em desenvolvimento, retornando ao menu........");
                        program = 0;
                        break;
                    case 9:
                        Console.WriteLine("Insira o internal name da lista: ");
                        InternalName = Console.ReadLine();
                        Console.WriteLine("Insira o nome do arquivo de destino (sem extenção): ");
                        file = Console.ReadLine();
                        Console.WriteLine("Insira o nome dos campos a serem exportados separados por vírgula, sem espaço: ");
                        Fields = Console.ReadLine();
                        operations.ExportList(InternalName, file, Fields);
                        break;

                    default:
                        program = 0;
                        break;
                        


                }




            }
        }
    }
}
