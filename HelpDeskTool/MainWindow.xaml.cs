﻿using HelpDeskTool;
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
    /// Created 9/25/2023
    /// Program HDTool written by Joseph Patrick for use in the URMC ISD Help Desk environment
    /// Credit to Alex McCune for initial LDAP usage and integration
    /// Contact: jpatrick5402@gmail.com
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
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Help Desk Tool Connection Error", (MessageBoxButtons)MessageBoxButton.OK, (MessageBoxIcon)MessageBoxImage.Error);
                Close(); return;
            }

            OutputBox.AppendText("\nHelp Desk Tool\n");
            OutputBox.AppendText("\nAwaiting Commands\n");
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void GroupMembersButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var GroupName = UserTextbox.Text.Trim();
                System.Windows.Clipboard.SetText(GroupName);
                UserTextbox.Clear();

                OutputBox.AppendText("Gathering Group Members for " + GroupName + "\n\n");

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=group)(name=" + GroupName + "))";
                SearchResult GroupMembersResult = searcher.FindOne();

                if (GroupMembersResult != null)
                {

                    string[] SortedGroup = new string[GroupMembersResult.Properties["member"].Count];
                    if (GroupMembersResult.Properties["member"].Count != 0)
                    {
                        for (int i = 0; i < GroupMembersResult.Properties["member"].Count; i++)
                        {
                            entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                            searcher = new DirectorySearcher(entry);

                            searcher.Filter = "(&(objectClass=user)(name=" + GroupMembersResult.Properties["member"][i].ToString().Substring(3, GroupMembersResult.Properties["member"][i].ToString().IndexOf(",OU") - 3).Replace("\\", "") + "))";
                            SearchResult GroupUserResult = searcher.FindOne();

                            SortedGroup[i] = GroupMembersResult.Properties["member"][i].ToString().Substring(3, GroupMembersResult.Properties["member"][i].ToString().IndexOf(",OU") - 3).Replace("\\", "") + "\t\t\t" + GroupUserResult.Properties["samaccountname"][0].ToString() + '\n';
                        }
                        Array.Sort(SortedGroup);
                        OutputBox.AppendText($"Members of {GroupMembersResult.Properties["name"][0]}:\n\n");
                        foreach (var item in SortedGroup)
                        {
                            OutputBox.AppendText(item + '\n');
                        }
                    }
                    else
                    {
                        string KeyString = "";
                        foreach (var item in GroupMembersResult.Properties.PropertyNames)
                        {
                            if (item.ToString().Contains("member;"))
                            {
                                KeyString = item.ToString();
                            }
                        }
                        if (KeyString != "")
                        {
                            for (int i = 0; i < GroupMembersResult.Properties[KeyString].Count; i++)
                            {
                                OutputBox.AppendText(GroupMembersResult.Properties[KeyString][i].ToString().Substring(3, GroupMembersResult.Properties[KeyString][i].ToString().IndexOf(",OU") - 3).Replace("\\", "") + '\n');
                            }
                        }
                        else
                        {
                            OutputBox.AppendText($"Unable to find members of group {GroupName}");
                        }
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find group \"{GroupName}\"");
                }
            }
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }

        private void UserInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var UserName = UserTextbox.Text.Trim();
                OutputBox.AppendText("Gathering info for " + UserName + "\n\n");
                System.Windows.Clipboard.SetText(UserName);
                UserTextbox.Clear();

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=user)(sAMAccountName=" + UserName + "))";
                SearchResult UserResult = searcher.FindOne();

                if (UserResult == null)
                {
                    searcher.Filter = "(&(objectClass=user)(urid=" + UserName + "))";
                    UserResult = searcher.FindOne();
                }
                if (UserResult == null)
                {
                    searcher.Filter = "(&(objectClass=user)(mail=" + UserName + "))";
                    UserResult = searcher.FindOne();
                }
                if (UserResult == null)
                {
                    searcher.Filter = "(&(objectClass=user)(name=" + UserName + "))";
                    UserResult = searcher.FindOne();
                }
                if (UserResult != null)
                {

                    OutputBox.AppendText("Gathering info for " + UserResult.Properties["name"][0] + " (" + UserResult.Properties["samaccountname"][0].ToString() + ")\n\n");

                    // Grabbing common items
                    string[,] PropertyList = { { "First Name", "givenname" }, { "Last Name", "sn" }, { "URMC AD", "samaccountname" }, { "UR AD", "" }, { "NetID", "uid" }, { "URID", "urid" }, { "Title", "title" }, { "Dept.", "department" }, { "Email", "mail" }, { "Phone", "telephoneNumber" }, { "Most Recent HR Role", "urrolestatus" }, { "Bad Password Count (Not Always Accurate)", "badpwdcount" } };

                    for (int i = 0; i < PropertyList.Length / 2; i++)
                    {
                        if (PropertyList[i, 0] == "UR AD")
                        {
                            // Grab UR AD username
                            try
                            {
                                DirectoryEntry URentry = new DirectoryEntry("LDAP://ur.rochester.edu");
                                DirectorySearcher URsearcher = new DirectorySearcher(entry);

                                searcher.Filter = $"(&(objectClass=user)(uidnumber={UserResult.Properties["uidnumber"]}))";
                                SearchResult URUserResult = searcher.FindOne();

                                if (URUserResult != null)
                                {
                                    OutputBox.AppendText("UR AD: " + URUserResult.Properties["uid"][0].ToString() + "\n");
                                }
                                else
                                {
                                    OutputBox.AppendText("UR AD: No active UR Active Directory account\n");
                                }
                            }
                            catch (Exception ex)
                            {
                                OutputBox.AppendText(ex.Message + "\n");
                            }
                        }
                        else if (PropertyList[i, 1] == "urrolestatus")
                        {
                            foreach (var item in UserResult.Properties[PropertyList[i, 1]])
                            {
                                OutputBox.AppendText("HR relationship: " + item.ToString() + "\n");
                            }
                        }
                        else
                        {
                            try
                            {
                                OutputBox.AppendText($"{PropertyList[i, 0]}: " + UserResult.Properties[PropertyList[i, 1]][0] + '\n');
                            }
                            catch
                            {
                                OutputBox.AppendText($"{PropertyList[i, 0]} is not listed in object properties\n");
                            }
                        }
                    }

                    using (PrincipalContext context = new PrincipalContext(ContextType.Domain, "urmc-sh.rochester.edu"))
                    {
                        /// Grabbing Pwd Last set. This has to use the PrincipalSearch module as DirectorySearch displays a value in "INTER8" format
                        /// which doesn't seem to have a 1:1 conversion to DateTime
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, UserResult.Properties["samaccountname"][0].ToString());
                        DateTime? passwordLastSet = user.LastPasswordSet;

                        if (passwordLastSet != null)
                        {
                            OutputBox.AppendText("Password Last Set: " + passwordLastSet.ToString() + '\n');
                            TimeSpan diff = DateTime.Today - passwordLastSet.Value;
                            if (diff.TotalDays >= 365)
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

                        // Grabbing OU
                        using (DirectoryEntry de = user.GetUnderlyingObject() as DirectoryEntry)
                        {
                            de.RefreshCache(new string[] { "canonicalName" });
                            string canonicalName = de.Properties["canonicalName"].Value as string;
                            OutputBox.AppendText($"OU: {canonicalName}\n");
                        }
                    }

                    /// This section is where we search for any shared mailboxes, this is done by searching each line of the files used for this task
                    /// and checking to see if the value of 'samaccountname' for our user is valid
                    try
                    {
                        OutputBox.AppendText("\n");
                        bool HasMailboxOrDL = false;
                        using (var sr = new StreamReader("\\\\nt014\\AdminApps\\Utils\\AD Utilities\\HDAMU-Support\\ResourceMailboxOwners.csv"))
                        {
                            string[] MailboxOwners = sr.ReadToEnd().Split('\n');

                            for (int i = 0; i < MailboxOwners.Length; i++)
                            {
                                if (MailboxOwners[i].Contains(UserResult.Properties["samaccountname"][0].ToString()))
                                {
                                    OutputBox.AppendText("Owned Mailbox: " + MailboxOwners[i].ToString().Substring(0, MailboxOwners[i].ToString().IndexOf(",")) + "\n");
                                    HasMailboxOrDL = true;
                                }
                            }
                        }
                        using (var sr = new StreamReader("\\\\nt014\\AdminApps\\Utils\\AD Utilities\\HDAMU-Support\\MigratedDistributionGroupExport.csv"))
                        {
                            string[] DLOwners = sr.ReadToEnd().Split('\n');

                            for (int i = 0; i < DLOwners.Length; i++)
                            {
                                if (DLOwners[i].Contains(UserResult.Properties["name"][0].ToString()))
                                {
                                    OutputBox.AppendText("Owned DL: " + DLOwners[i].ToString().Substring(0, DLOwners[i].ToString().IndexOf(",")) + "\n");
                                    HasMailboxOrDL = true;
                                }
                            }
                        }
                        using (var sr = new StreamReader("\\\\nt014\\AdminApps\\Utils\\AD Utilities\\HDAMU-Support\\Mailbox-Owners-Managers.csv"))
                        {
                            // Information from here may not be the most accurate as it's not the exact same as HDAMU
                            string[] MailboxOwners = sr.ReadToEnd().Split('\n');


                            for (int i = 0; i < MailboxOwners.Length; i++)
                            {
                                if (MailboxOwners[i].Contains(UserResult.Properties["mail"][0].ToString()))
                                {
                                    OutputBox.AppendText("Owned Mailbox (Mailbox may not be active): " + MailboxOwners[i].ToString().Substring(0, MailboxOwners[i].ToString().IndexOf(",")) + "\n");
                                    HasMailboxOrDL = true;
                                }
                            }
                        }
                        if (!HasMailboxOrDL)
                        {
                            OutputBox.AppendText(UserResult.Properties["name"][0].ToString() + " owns no DLs or Shared Mailboxes\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText(ex.Message);
                        OutputBox.AppendText("Unable to fetch Shared Mailboxes/DLs\n");
                    }

                    /// In this section, we take a look at the shared drive files pulled directly from the source DMD files and check to see if our user has
                    /// access to any of them. We have the potential to also add the AD group that grants access, but that may cause bloat to the output.
                    try
                    {
                        OutputBox.AppendText("\nShared/Home Drive Access List: \n");
                        entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                        searcher = new DirectorySearcher(entry);

                        searcher.Filter = "(&(objectClass=*)(sAMAccountName=" + UserResult.Properties["samaccountname"][0].ToString() + "))";
                        SearchResult MemberOfresult = searcher.FindOne();

                        if (MemberOfresult != null)
                        {
                            OutputBox.AppendText(MemberOfresult.Properties["homedirectory"][0].ToString() + "\n");

                            ResultPropertyValueCollection groups = MemberOfresult.Properties["memberOf"];
                            using (var sr = new StreamReader("\\\\ADSDC01\\netlogon\\SIG\\logon.dmd"))
                            {
                                string[] ShareList = sr.ReadToEnd().Split("\n");
                                foreach (var group in groups)
                                {
                                    for (int i = 0; i < ShareList.Length; i++)
                                    {
                                        if (ShareList[i].Contains(group.ToString().Substring(3, group.ToString().IndexOf(",") - 3) + "|"))
                                        {
                                            OutputBox.AppendText(ShareList[i].Substring(ShareList[i].LastIndexOf('|') + 1, ShareList[i].Substring(ShareList[i].LastIndexOf('|')).Length - 2) + " | " + ShareList[i].Substring(0, ShareList[i].IndexOf('|')) + '\n');
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText(ex.Message);
                        OutputBox.AppendText($"Unable to find Shared Drives for {UserName}");
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find username \"{UserName}\"");
                }

            }
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void ComputerInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var ComputerName = NameBox.Text.Trim();
                OutputBox.AppendText("Gathering info for " + ComputerName + "\n\n");
                System.Windows.Clipboard.SetText(ComputerName);
                NameBox.Clear();

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = $"(&(objectClass=computer)(CN={ComputerName}))";
                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    string[,] PropertyList = { { "Domain Name", "DNSHostName" }, { "OS", "operatingsystem" }, { "OS Version", "operatingsystemversion" } };

                    for (int i = 0; i < PropertyList.Length / 2; i++)
                    {
                        try
                        {
                            OutputBox.AppendText($"{PropertyList[i, 0]}: " + result.Properties[PropertyList[i, 1]][0] + '\n');
                        }
                        catch
                        {
                            OutputBox.AppendText($"{PropertyList[i, 0]} is not listed on properties\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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

                    string[,] PropertyList = { { "Name", "cn" }, { "Last Change Date", "whenchanged" }, { "Creation Date", "whencreated" }, { "Description", "description" }, { "Additional Info", "info" }, { "Manager", "managedby" } };

                    for (int i = 0; i < PropertyList.Length / 2; i++)
                    {
                        try
                        {
                            OutputBox.AppendText($"{PropertyList[i, 0]}: " + GroupResultInfo.Properties[PropertyList[i, 1]][0] + '\n');
                        }
                        catch
                        {
                            OutputBox.AppendText($"{PropertyList[i, 0]} is not listed in object properties\n");
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
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
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
            string helpInfo = "HDTool is an AD/computer management tool to improve the efficiency of the Help Desk" +
                "\n" +
                "\nOnce information is entered in the \"AD Name\" or the \"PC Name/IP\" boxes, you can click on any button next to that input box to perform action on that item" +
                "\nPlease reach out to Joseph Patrick on Teams or joseph_patrick@urmc.rochester.edu / jpatrick5402@gmail.com for any questions comments or concerns";
            OutputBox.AppendText(helpInfo);
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }

        private void MasterSearchButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var SearchObject = UserTextbox.Text.Trim();
                OutputBox.AppendText("Searching for " + SearchObject + "\n\n");
                System.Windows.Clipboard.SetText(SearchObject);
                UserTextbox.Clear();

                // Search under URMC umbrella
                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = $"(&(objectClass=*)(cn={SearchObject}*))";
                SearchResultCollection Result = searcher.FindAll();
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(urid={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(mail={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(samaccountname={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(name={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    OutputBox.AppendText("URMC: No object found\n");
                }
                else
                {

                    foreach (SearchResult result in Result)
                    {
                        OutputBox.AppendText("URMC: " + result.Properties["cn"][0].ToString() + "\t\t" + result.Properties["objectclass"][^1].ToString() + '\n');
                    }
                }

                // Search under UR umbrella
                entry = new DirectoryEntry("LDAP://ur.rochester.edu");
                searcher = new DirectorySearcher(entry);

                searcher.Filter = $"(&(objectClass=*)(cn={SearchObject}))";
                Result = searcher.FindAll();

                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(urid={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(mail={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(samaccountname={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    searcher.Filter = $"(&(objectClass=*)(name={SearchObject}*))";
                    Result = searcher.FindAll();
                }
                if (Result.Count == 0)
                {
                    OutputBox.AppendText("UR: No object found\n");
                }
                else
                {

                    foreach (SearchResult result in Result)
                    {
                        OutputBox.AppendText("UR: " + result.Properties["cn"][0].ToString() + "\t\t" + result.Properties["objectclass"][^1].ToString() + '\n');
                    }
                }
            }
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }

        private void DarkButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindow1.Background = new LinearGradientBrush(Colors.Navy, Colors.Black, 90.00);
            OutputBox.Background = Brushes.DarkGray;
            PCNameBox.Foreground = Brushes.White;
            ADNameBox.Foreground = Brushes.White;
            NameBox.Background = Brushes.DarkGray;
            UserTextbox.Background = Brushes.DarkGray;
        }
    }
}