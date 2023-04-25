using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.IO;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.News.DataModel;
using System.Xml;
using System.Xml.Linq;

namespace The_SharePoint_Machine
{
    internal class SP_Operations 
    {
        private ClientContext context;

        public SP_Operations(ClientContext context)
        {
            this.context = context;
        }

        public void CreateSite(string InternalName, string Title, string Description)
        {
            try
            {
                WebCreationInformation creation = new WebCreationInformation();

                creation.Url = InternalName;
                creation.Title = Title;
                creation.Description = Description;

                Web newWeb = context.Web.Webs.Add(creation);

                this.context.Load(newWeb);
                this.context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }

        }

        public void DeleteSite()
        {
            try
            {
                Web web = context.Web;
                web.DeleteObject();
                context.ExecuteQuery();

            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }

        }

        public void CreateRole(string name, string description, string[] permissions)
        {
            try
            {
                BasePermissions perm = new BasePermissions();
                foreach (string permission in permissions)
                {
                    switch (Int32.Parse(permission))
                    {
                        case 1:
                            perm.Set(PermissionKind.ViewListItems);
                            break;
                        case 2:
                            perm.Set(PermissionKind.AddListItems);
                            break;
                        case 3:
                            perm.Set(PermissionKind.EditListItems);
                            break;
                        case 4: 
                            perm.Set(PermissionKind.DeleteListItems);
                            break;
                        case 5:
                            perm.Set(PermissionKind.ApproveItems);
                            break;
                        case 6:
                            perm.Set(PermissionKind.OpenItems);
                            break;
                        case 7:
                            perm.Set(PermissionKind.ViewVersions);
                            break;
                        case 8:
                            perm.Set(PermissionKind.DeleteVersions);
                            break;
                        case 9:
                            perm.Set(PermissionKind.CancelCheckout);
                            break;
                        case 10:
                            perm.Set(PermissionKind.ManagePersonalViews);
                            break;
                        case 11:
                            perm.Set(PermissionKind.ManageLists);
                            break;
                        case 12:
                            perm.Set(PermissionKind.ViewFormPages);
                            break;
                        case 13:
                            perm.Set(PermissionKind.CreateAlerts);
                            break;
                        case 14:
                            perm.Set(PermissionKind.ManageAlerts);
                            break;
                        case 15:
                            perm.Set(PermissionKind.ViewPages);
                            break;
                        case 16:
                            perm.Set(PermissionKind.AddAndCustomizePages);
                            break;
                        case 18:
                            perm.Set(PermissionKind.ApplyThemeAndBorder);
                            break;
                        case 19:
                            perm.Set(PermissionKind.ApplyStyleSheets);
                            break;
                        case 20:
                            perm.Set(PermissionKind.ViewUsageData);
                            break;
                        case 21:
                            perm.Set(PermissionKind.ManageSubwebs);
                            break;
                        case 22:
                            perm.Set(PermissionKind.ManagePermissions);
                            break;
                        case 23:
                            perm.Set(PermissionKind.BrowseDirectories);
                            break;
                        case 24:
                            perm.Set(PermissionKind.EditMyUserInfo);
                            break;
                        case 25:
                            perm.Set(PermissionKind.UseClientIntegration);
                            break;
                        case 26:
                            perm.Set(PermissionKind.UseRemoteAPIs);
                            break;
                        case 27:
                            perm.Set(PermissionKind.EnumeratePermissions);
                            break;
                        case 28:
                            perm.Set(PermissionKind.FullMask);
                            break;

                    }
                }
                RoleDefinitionCreationInformation creatInfo = new RoleDefinitionCreationInformation();
                creatInfo.BasePermissions = perm;
                creatInfo.Name = name;
                creatInfo.Description = description;

                RoleDefinition rd = context.Web.RoleDefinitions.Add(creatInfo);
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void EditRole(string name, string description, string[] Addpermissions, string[] RemovePermissions)
        {
            try
            {
                Web web = context.Web;

                BasePermissions perm = new BasePermissions();
                foreach (string permission in Addpermissions)
                {
                    switch (Int32.Parse(permission))
                    {
                        case 1:
                            perm.Set(PermissionKind.ViewListItems);
                            break;
                        case 2:
                            perm.Set(PermissionKind.AddListItems);
                            break;
                        case 3:
                            perm.Set(PermissionKind.EditListItems);
                            break;
                        case 4:
                            perm.Set(PermissionKind.DeleteListItems);
                            break;
                        case 5:
                            perm.Set(PermissionKind.ApproveItems);
                            break;
                        case 6:
                            perm.Set(PermissionKind.OpenItems);
                            break;
                        case 7:
                            perm.Set(PermissionKind.ViewVersions);
                            break;
                        case 8:
                            perm.Set(PermissionKind.DeleteVersions);
                            break;
                        case 9:
                            perm.Set(PermissionKind.CancelCheckout);
                            break;
                        case 10:
                            perm.Set(PermissionKind.ManagePersonalViews);
                            break;
                        case 11:
                            perm.Set(PermissionKind.ManageLists);
                            break;
                        case 12:
                            perm.Set(PermissionKind.ViewFormPages);
                            break;
                        case 13:
                            perm.Set(PermissionKind.CreateAlerts);
                            break;
                        case 14:
                            perm.Set(PermissionKind.ManageAlerts);
                            break;
                        case 15:
                            perm.Set(PermissionKind.ViewPages);
                            break;
                        case 16:
                            perm.Set(PermissionKind.AddAndCustomizePages);
                            break;
                        case 18:
                            perm.Set(PermissionKind.ApplyThemeAndBorder);
                            break;
                        case 19:
                            perm.Set(PermissionKind.ApplyStyleSheets);
                            break;
                        case 20:
                            perm.Set(PermissionKind.ViewUsageData);
                            break;
                        case 21:
                            perm.Set(PermissionKind.ManageSubwebs);
                            break;
                        case 22:
                            perm.Set(PermissionKind.ManagePermissions);
                            break;
                        case 23:
                            perm.Set(PermissionKind.BrowseDirectories);
                            break;
                        case 24:
                            perm.Set(PermissionKind.EditMyUserInfo);
                            break;
                        case 25:
                            perm.Set(PermissionKind.UseClientIntegration);
                            break;
                        case 26:
                            perm.Set(PermissionKind.UseRemoteAPIs);
                            break;
                        case 27:
                            perm.Set(PermissionKind.EnumeratePermissions);
                            break;
                        case 28:
                            perm.Set(PermissionKind.FullMask);
                            break;

                    }
                }
                foreach (string permission in RemovePermissions)
                {
                    switch (Int32.Parse(permission))
                    {
                        case 1:
                            perm.Clear(PermissionKind.ViewListItems);
                            break;
                        case 2:
                            perm.Clear(PermissionKind.AddListItems);
                            break;
                        case 3:
                            perm.Clear(PermissionKind.EditListItems);
                            break;
                        case 4:
                            perm.Clear(PermissionKind.DeleteListItems);
                            break;
                        case 5:
                            perm.Clear(PermissionKind.ApproveItems);
                            break;
                        case 6:
                            perm.Clear(PermissionKind.OpenItems);
                            break;
                        case 7:
                            perm.Clear(PermissionKind.ViewVersions);
                            break;
                        case 8:
                            perm.Clear(PermissionKind.DeleteVersions);
                            break;
                        case 9:
                            perm.Clear(PermissionKind.CancelCheckout);
                            break;
                        case 10:
                            perm.Clear(PermissionKind.ManagePersonalViews);
                            break;
                        case 11:
                            perm.Clear(PermissionKind.ManageLists);
                            break;
                        case 12:
                            perm.Clear(PermissionKind.ViewFormPages);
                            break;
                        case 13:
                            perm.Clear(PermissionKind.CreateAlerts);
                            break;
                        case 14:
                            perm.Clear(PermissionKind.ManageAlerts);
                            break;
                        case 15:
                            perm.Clear(PermissionKind.ViewPages);
                            break;
                        case 16:
                            perm.Clear(PermissionKind.AddAndCustomizePages);
                            break;
                        case 18:
                            perm.Clear(PermissionKind.ApplyThemeAndBorder);
                            break;
                        case 19:
                            perm.Clear(PermissionKind.ApplyStyleSheets);
                            break;
                        case 20:
                            perm.Clear(PermissionKind.ViewUsageData);
                            break;
                        case 21:
                            perm.Clear(PermissionKind.ManageSubwebs);
                            break;
                        case 22:
                            perm.Clear(PermissionKind.ManagePermissions);
                            break;
                        case 23:
                            perm.Clear(PermissionKind.BrowseDirectories);
                            break;
                        case 24:
                            perm.Clear(PermissionKind.EditMyUserInfo);
                            break;
                        case 25:
                            perm.Clear(PermissionKind.UseClientIntegration);
                            break;
                        case 26:
                            perm.Clear(PermissionKind.UseRemoteAPIs);
                            break;
                        case 27:
                            perm.Clear(PermissionKind.EnumeratePermissions);
                            break;
                        case 28:
                            perm.Clear(PermissionKind.FullMask);
                            break;

                    }
                }
                RoleDefinition UpdateInfo = web.RoleDefinitions.GetByName(name);
                UpdateInfo.BasePermissions = perm;
                if (description != null)
                {
                    UpdateInfo.Description = description;
                }

                UpdateInfo.Update();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }

        }

        public void DeleteRole (string Name)
        {
            try
            {
                Web web = context.Web;
                RoleDefinition role = web.RoleDefinitions.GetByName(Name);

                // Delete Role
                role.DeleteObject();

                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }

        }

        public void AddUsersToGroup(string groupName, string[] users)
        {
            try
            {
                Web web = context.Web;
                Group group = web.SiteGroups.GetByName(groupName);

                foreach (string item in users)
                {
                    User user = web.EnsureUser(item);
                    group.Users.AddUser(user);
                }
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }

        }

        public void RemoveUsersOfGroup(string groupName, string[] users)
        {
            try
            {
                Web web = context.Web;
                Group group = web.SiteGroups.GetByName(groupName);

                foreach (string item in users)
                {
                    User user = web.EnsureUser(item);
                    group.Users.RemoveByLoginName(user.LoginName);
                }
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }

        }
        public void AddPermitionToGroup(string groupName, string permissionName)
        {
            try
            {
                Web web = context.Web;
                RoleDefinitionCollection roleDefinitions = web.RoleDefinitions;
                RoleDefinition roleDefinition = roleDefinitions.GetByName(permissionName);
                Group group = web.SiteGroups.GetByName(groupName);
                var roleAssignment = new RoleDefinitionBindingCollection(context) { roleDefinition };
                web.RoleAssignments.Add(group, roleAssignment);
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void RemovePermitionOfGroup(string groupName, string permissionName)
        {
            try
            {
                Web web = context.Web;
                RoleDefinitionCollection roleDefinitions = web.RoleDefinitions;
                RoleDefinition roleDefinition = roleDefinitions.GetByName(permissionName);
                Group group = web.SiteGroups.GetByName(groupName);
                RoleAssignmentCollection roleAssignments = web.RoleAssignments;
                RoleAssignment roleAssignment = roleAssignments.GetByPrincipal(group);
                roleAssignment.RoleDefinitionBindings.Remove(roleDefinition);
                roleAssignment.Update();
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void CreateGroup(string name, string descrription, string[] users, string[] permissions)
        {
            try
            {
                Web web = context.Web;

                GroupCreationInformation createGroup = new GroupCreationInformation();
                createGroup.Title = name;
                createGroup.Description = descrription;
                web.SiteGroups.Add(createGroup);
                context.ExecuteQuery();

                if (users != null)
                {
                    this.AddUsersToGroup(name, users);
                }
                if (permissions != null)
                {
                    foreach (string permission in permissions)
                    {
                        this.AddPermitionToGroup(name, permission);
                    }
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void EditGroup(string groupName, string newGroupName, string description, string[] addedUsers, string[] removeUser, string[] removePermission, string[] addPermission)
        {
            try
            {
                Web web = context.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName(groupName);
                if(description != null)
                {
                    group.Description = description;
                }
                if(addedUsers != null)
                {
                    this.AddUsersToGroup(groupName, addedUsers);
                }
                if (removeUser != null)
                {
                    this.RemoveUsersOfGroup(groupName, removeUser);
                }
                if (removePermission != null)
                {
                    foreach (string permission in removePermission)
                    {
                        this.RemovePermitionOfGroup(groupName, permission);
                    }
                }
                if (addPermission != null)
                {
                    foreach (string permission in addPermission)
                    {
                        this.RemovePermitionOfGroup(groupName, permission);
                    }
                }

                if(newGroupName !=  null){
                    group.Title = newGroupName;
                }
                group.Update();
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void DeleteGroup(string groupName)
        {
            try
            {
                Web web = context.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName(groupName);
                groups.Remove(group);
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void CreateFields(string XmlPath, string ListName)
        {
            try
            {
                List<string> fieldsXml = new List<string>();
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                using (XmlReader reader = XmlReader.Create(XmlPath))
                {
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "Field")
                        {
                            fieldsXml.Add(reader.ReadOuterXml());
                        }
                    }
                }
                foreach (string fieldXml in fieldsXml)
                {

                    list.Fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.DefaultValue);
                }
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }


        }
        public void CreateList(string name, string internalName, string description, string XmlPath)
        {
            try
            {
                Web web = context.Web;
                ListCreationInformation createList = new ListCreationInformation();
                createList.Title = name;
                createList.TemplateType = (int)ListTemplateType.GenericList;
                createList.Description = description;
                createList.Url = internalName;
                List list = web.Lists.Add(createList);
                context.ExecuteQuery();

            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Ocorreu um erro: {e}");
                Console.ResetColor();
            }
            this.CreateFields(XmlPath, internalName);
        }

        public void EditList(string internalName, string xml, string[] fields)
        {
            try
            {
                Web web = context.Web;
                List list = web.Lists.GetByTitle(internalName);
                if(xml != null)
                {
                    list.Fields.AddFieldAsXml(xml, true, AddFieldOptions.DefaultValue);
                }
                if(fields != null)
                {
                    foreach (string fieldName in fields)
                    {
                        Field field = list.Fields.GetByInternalNameOrTitle(fieldName);
                        field.DeleteObject();
                    }
                }
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro:" + e.Message);
                Console.ResetColor();
            }
        }

        public void DeleteList(string InternalName)
        {
            Web web = context.Web;
            List list = web.Lists.GetByTitle(InternalName);
            list.DeleteObject();
            context.ExecuteQuery();
        }

        public void ExportList(string internalName, string fileName, string fields)
        {
            List list = context.Web.Lists.GetByTitle(internalName);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = list.GetItems(query);

            context.Load(items);
            context.ExecuteQuery();

            StringBuilder csv = new StringBuilder();
            csv.Append(string.Join(",", fields.Replace(";", ".,") + ";"));
            foreach (ListItem item in items)
            {
                List<string> row = new List<string>();
                
                foreach (var field in fields)
                {
                    Console.WriteLine(field);
                }
                
            }

            
        }




    }
}
