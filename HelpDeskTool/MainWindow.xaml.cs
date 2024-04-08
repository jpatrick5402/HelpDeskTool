﻿/// Program HDTool written by Joseph Patrick for use in the URMC ISD Help Desk environment
/// Credit to Alex McCune for initial LDAP usage and integration
/// Created 9/25/2023

using HelpDeskTool;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.DirectoryServices;
using System.Windows.Forms;
using System.Collections.Immutable;
using System.Data;
using System.DirectoryServices.AccountManagement;

namespace DTTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            try
            {
                OutputBox.AppendText("\rConnecting to LDAP server...\n");
                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);
                searcher.Filter = "(&(objectClass=user)(sAMAccountName=" + Environment.UserName + "))";
                SearchResult result = searcher.FindOne();
                if (result == null) 
                {
                    throw new InvalidOperationException("error");
                }
                OutputBox.AppendText("Connected\n");
            }
            catch (Exception e) {
                System.Windows.Forms.MessageBox.Show(e.Message, "Help Desk Tool Connection Error", (MessageBoxButtons)MessageBoxButton.OK, (MessageBoxIcon)MessageBoxImage.Error);
                Close(); return;
            }

            OutputBox.AppendText("\nHelp Desk Tool\n");
            OutputBox.AppendText("\nAwaiting Commands\n");
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
        }
        public Boolean IsTextInNameBox()
        {
            if (NameBox.Text != "")
                return true;
            System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            OutputBox.AppendText("No PC Name/IP Detected");
            return false;
        }
        public Boolean IsTextInUserBox()
        {
            if (UserTextbox.Text != "")
                return true;
            System.Windows.MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            OutputBox.AppendText("No AD Name Detected");
            return false;
        }
        private void RestartButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text.Trim();

                MessageBoxResult result = System.Windows.MessageBox.Show("Are you sure you want to restart this computer? " + PCName,
                                          "Confirmation",
                                          MessageBoxButton.YesNo,
                                          MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    System.Windows.Clipboard.SetText(PCName);
                    NameBox.Clear();
                    System.Diagnostics.Process command = new System.Diagnostics.Process();
                    command.StartInfo.CreateNoWindow = true;
                    command.StartInfo.FileName = "cmd";
                    command.StartInfo.Arguments = "/C shutdown -r -t 2 -m " + PCName;
                    command.StartInfo.RedirectStandardOutput = true;
                    command.Start();
                    OutputBox.AppendText("Restart Initiated (Ask to ensure as this is not 100% accurate)");
                    OutputBox.AppendText(command.StandardOutput.ReadToEnd());
                }
                else
                {
                    OutputBox.AppendText("PC not restarted\n");
                }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void PingButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text.Trim();
                OutputBox.AppendText("Pinging " + PCName + "\n\n");

                System.Windows.Clipboard.SetText(PCName);
                NameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                command.StartInfo.Arguments = "/C ping /n 2 " + PCName;
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void NslookupButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text.Trim();
                OutputBox.AppendText("Looking for " + PCName + "\n\n");

                System.Windows.Clipboard.SetText(PCName);
                NameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                command.StartInfo.Arguments = "/C nslookup " + PCName;
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void SysinfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text.Trim();
                OutputBox.AppendText("Gathering info on " + PCName + "\n\n");

                System.Windows.Clipboard.SetText(PCName);
                NameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                command.StartInfo.Arguments = "/C powershell systeminfo /S " + PCName;
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void RemoteButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text.Trim();
                System.Windows.Clipboard.SetText(PCName);
                NameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "C:\\Program Files (x86)\\SCCM Remote Control\\CmRcViewer.exe";
                command.StartInfo.Arguments = PCName;
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText("Remote Session Initiated with " + PCName);
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void UserInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var UserName = UserTextbox.Text.Trim();

                System.Windows.Clipboard.SetText(UserName);
                UserTextbox.Clear();

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=user)(sAMAccountName=" + UserName + "))";
                SearchResult UserResult = searcher.FindOne();

                if (UserResult != null)
                {

                    OutputBox.AppendText("Gathering info for " + UserResult.Properties["name"][0] + " (" + UserName + ")\n\n");

                    string[] PropertyList = { "givenname", "sn", "samaccountname", "urid", "department", "mail", "telephoneNumber", "urrolestatus", "badpwdcount" };

                    foreach (string Property in PropertyList)
                    {
                        try
                        {
                            OutputBox.AppendText($"{Property}: " + UserResult.Properties[Property][0] + '\n');
                        }
                        catch
                        {
                            OutputBox.AppendText($"{Property} is not listed in object properties\n");
                        }
                    }

                    using (PrincipalContext context = new PrincipalContext(ContextType.Domain, "urmc-sh.rochester.edu"))
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, UserName);
                        DateTime? passwordLastSet = user.LastPasswordSet;

                        if (passwordLastSet != null)
                        {
                            OutputBox.AppendText("Password Last Set: " + passwordLastSet.ToString() + '\n');
                            TimeSpan diff = passwordLastSet.Value - DateTime.Today;
                            if ((diff).TotalDays < -365)
                            {
                                OutputBox.AppendText("Pwd Expired: True\n");
                            }
                            else
                            {
                                OutputBox.AppendText("Pwd Expired: False\n");
                            }
                        }
                        else
                        {
                            OutputBox.AppendText("Password Last Set information is not available.\n");
                        }
                        using (DirectoryEntry de = user.GetUnderlyingObject() as DirectoryEntry)
                        {
                            de.RefreshCache(new string[] { "canonicalName" });
                            string canonicalName = de.Properties["canonicalName"].Value as string;
                            OutputBox.AppendText($"OU: {canonicalName}");
                        }
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find username \"{UserName}\"");
                }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void MemberOfButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                
                var UserName = UserTextbox.Text.Trim();
                System.Windows.Clipboard.SetText(UserName);
                UserTextbox.Clear();

                OutputBox.AppendText("Gathering Memberships for " + UserName + "\n\n");

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=*)(sAMAccountName=" + UserName + "))";
                SearchResult MemberOfresult = searcher.FindOne();

                if (MemberOfresult != null)
                {
                    ResultPropertyValueCollection groups = MemberOfresult.Properties["memberOf"];
                    OutputBox.AppendText($"{MemberOfresult.Properties["name"][0]} ({MemberOfresult.Properties["SAMAccountName"][0]}) is a member of {groups.Count} groups\n\n");
                    if (groups.Count > 0)
                    {
                        string[] memberships = new string[groups.Count];
                        for (int i = 0; i < groups.Count; i++)
                        {
                            memberships[i] = groups[i].ToString();
                            memberships[i] = memberships[i].ToString().Substring(3, memberships[i].IndexOf(",") - 3);
                        }
                        Array.Sort(memberships);
                        foreach (var group in memberships)
                        {
                            OutputBox.AppendText(group + '\n');
                        }
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find object \"{UserName}\"");
                }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void ComputerInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var ComputerName = NameBox.Text.Trim();
                OutputBox.AppendText("Gathering Computer's AD info for " + ComputerName + "\n\n");
                System.Windows.Clipboard.SetText(ComputerName);

                NameBox.Clear();
                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = $"(&(objectClass=computer)(CN={ComputerName}))";
                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    string[] PropertyList = { "cn", "operatingsystem", "operatingsystemversion", "ms-mcs-admpwd" };

                    foreach (string Property in PropertyList)
                    {
                        try
                        {
                            OutputBox.AppendText($"{Property}: " + result.Properties[Property][0] + '\n');
                        }
                        catch
                        {
                            OutputBox.AppendText($"{Property} is not listed on properties\n");
                        }
                    }
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "urmc-sh.rochester.edu");
                    ComputerPrincipal computer = ComputerPrincipal.FindByIdentity(ctx, ComputerName);
                    using (DirectoryEntry de = computer.GetUnderlyingObject() as DirectoryEntry)
                    {
                        de.RefreshCache(new string[] { "canonicalName" });
                        string canonicalName = de.Properties["canonicalName"].Value as string;
                        OutputBox.AppendText("OU: " + canonicalName);
                    }
                }
                else
                { OutputBox.AppendText($"{ComputerName} not found"); }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void GroupMembersButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var GroupName = UserTextbox.Text.Trim();
                OutputBox.AppendText("Gathering Group Members for " + GroupName + "\n\n");

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=group)(CN=" + GroupName + "))";
                SearchResult GroupMembersResult = searcher.FindOne();

                if (GroupMembersResult != null)
                {

                    string[] SortedGroup = new string[GroupMembersResult.Properties["member"].Count];
                    int i = 0;
                    foreach (var item in GroupMembersResult.Properties["member"])
                    {
                        SortedGroup[i] = item.ToString().Substring(3, item.ToString().IndexOf(",OU") - 3).Replace("\\", "");
                        i++;
                    }
                    Array.Sort(SortedGroup);
                    OutputBox.AppendText($"Members of {GroupName}:\n");
                    foreach (var item in SortedGroup)
                    {
                        OutputBox.AppendText(item + '\n');
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find group \"{GroupName}\"");
                }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void GroupInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var GroupName = UserTextbox.Text.Trim();
                OutputBox.AppendText("Gathering info for " + GroupName + "\n\n");

                System.Windows.Clipboard.SetText(GroupName);
                UserTextbox.Clear();

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=group)(CN=" + GroupName + "))";
                SearchResult GroupResultInfo = searcher.FindOne();

                if (GroupResultInfo != null)
                {

                    string[] PropertyList = { "cn", "whenchanged", "whencreated", "description", "info", "managedby" };

                    foreach (string Property in PropertyList)
                    {
                        try
                        {
                            OutputBox.AppendText($"{Property}: " + GroupResultInfo.Properties[Property][0] + '\n');
                        }
                        catch
                        {
                            OutputBox.AppendText($"\"{Property}\" is not listed in object properties\n");
                        }
                    }
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "urmc-sh.rochester.edu");
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, GroupName);
                    using (DirectoryEntry de = group.GetUnderlyingObject() as DirectoryEntry)
                    {
                        de.RefreshCache(new string[] { "canonicalName" });
                        string canonicalName = de.Properties["canonicalName"].Value as string;
                        OutputBox.AppendText("OU: " + canonicalName);
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find group \"{GroupName}\"");
                }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void AddgroupButton_Click(object sender, RoutedEventArgs e)
        {
            AddRemoveWindow win = new AddRemoveWindow();
            win.Show();
        }
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            OutputBox.Document.Blocks.Clear();
        }
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            string helpInfo = "HDTool is an AD/computer management tool to improve the efficiency of the Help Desk\n\nOnce information is entered in the \"AD Name\" or the \"PC Name/IP\" boxes, you can click on any button next to that input to perform action on that item\n";
            OutputBox.AppendText(helpInfo);
            OutputBox.AppendText("---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
    }
}