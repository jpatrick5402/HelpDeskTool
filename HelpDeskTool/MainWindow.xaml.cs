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
using System.Net.Http;
using Microsoft.VisualBasic.ApplicationServices;
using System.Configuration;
using Microsoft.VisualBasic;
using System.Windows.Forms.PropertyGridInternal;

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
        // Initialization Function
        public MainWindow()
        {
            InitializeComponent();

            if (Settings.Default.DarkMode)
                DarkButton_Click(null, null);

            OutputBox.AppendText("\nHelp Desk Tool\n");
            OutputBox.AppendText("\nAwaiting Commands\n");
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
        }

        // Helper Functions
        public LoadingWindow ShowLoadingWindow()
        {
            LoadingWindow win = new LoadingWindow();
            win.Show();
            win.Topmost = true;
            win.Focus();
            return win;
        }
        public void CloseLoadingWindow(LoadingWindow win)
        {
            win.Close();
        }

        // Button Functions
        private void RestartButton_Click(object sender, RoutedEventArgs e)
        {
            if (NameBox.Text != "")
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void PingButton_Click(object sender, RoutedEventArgs e)
        {
            if (NameBox.Text != "")
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }
        private void NslookupButton_Click(object sender, RoutedEventArgs e)
        {
            if (NameBox.Text != "")
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void SysinfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (NameBox.Text != "")
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void RemoteButton_Click(object sender, RoutedEventArgs e)
        {
            if (NameBox.Text != "")
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void MemberOfButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserTextbox.Text != "")
            {

                var UserName = UserTextbox.Text.Trim();
                System.Windows.Clipboard.SetText(UserName);
                UserTextbox.Clear();

                OutputBox.AppendText("Gathering Memberships for " + UserName + "\n\n");

                DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu/DC=urmc-sh,DC=rochester,DC=edu");
                DirectorySearcher searcher = new DirectorySearcher(entry);

                searcher.Filter = "(&(objectClass=*)(sAMAccountName=" + UserName + "))";
                SearchResult MemberOfresult = searcher.FindOne();

                if (MemberOfresult == null)
                {
                    searcher.Filter = "(&(objectClass=*)(urid=" + UserName + "))";
                    MemberOfresult = searcher.FindOne();
                }
                if (MemberOfresult == null)
                {
                    searcher.Filter = "(&(objectClass=*)(name=" + UserName + "))";
                    MemberOfresult = searcher.FindOne();
                }
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void GroupMembersButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserTextbox.Text != "")
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

                            SortedGroup[i] = GroupMembersResult.Properties["member"][i].ToString().Substring(3, GroupMembersResult.Properties["member"][i].ToString().IndexOf(",OU") - 3).Replace("\\", "") + "\t" + GroupUserResult.Properties["samaccountname"][0].ToString();
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void UserInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserTextbox.Text != "")
            {
                var UserName = UserTextbox.Text.Trim();
                OutputBox.AppendText("Searching for " + UserName + "...\n");
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
                    string[,] PropertyList = { { "First Name:", "givenname" }, { "Last Name:", "sn" }, { "URMC AD:", "samaccountname" }, { "UR AD:\t", ""}, { "NetID:\t", "uid" }, { "URID:\t", "urid" }, { "Title:\t", "title" }, { "Dept.:\t", "department" }, { "Email:\t", "mail" }, { "Phone:\t", "telephoneNumber" }, { "Description:", "description" },  { "space", "" }, { "Most Recent HR Role", "urrolestatus" },{ "space", "" }, { "PWD Last Set:\t", "pwdlastset" }, { "OU", "adspath" }};

                    for (int i = 0; i < PropertyList.Length / 2; i++)
                    {
                        if (PropertyList[i, 0] == "space")
                        {
                            OutputBox.AppendText("\n");
                        }
                        else if (PropertyList[i, 0] == "Most Recent HR Role")
                        {
                            foreach (var item in UserResult.Properties[PropertyList[i, 1]])
                            {
                                OutputBox.AppendText("HR status:\t" + item.ToString() + "\n");
                            }
                        }
                        else if (PropertyList[i, 0].Contains("PWD Last Set"))
                        {
                            var pwdData = UserResult.Properties["pwdlastset"][0];
                            DateTime UnZonedDate = new DateTime(1601, 01, 01, 0, 0, 0, DateTimeKind.Utc).AddTicks((long)pwdData);
                            var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                            DateTime passwordLastSet = TimeZoneInfo.ConvertTime((DateTime)UnZonedDate, timeZone);

                            if (!timeZone.IsDaylightSavingTime(passwordLastSet))
                            {
                                passwordLastSet = passwordLastSet.AddHours(1);
                            }

                            OutputBox.AppendText(PropertyList[i,0] + passwordLastSet + "\n");

                            TimeSpan diff = DateTime.Today - passwordLastSet;
                            if (diff.TotalDays >= 365)
                            {
                                OutputBox.AppendText("Pwd Expired:\tTrue\n");
                            }
                            else
                            {
                                OutputBox.AppendText("Pwd Expired:\tFalse\n");
                            }
                        }
                        else if (PropertyList[i, 0] == "OU")
                        {
                            using (PrincipalContext context = new PrincipalContext(ContextType.Domain, "urmc-sh.rochester.edu"))
                            {
                                UserPrincipal user = UserPrincipal.FindByIdentity(context, UserResult.Properties["samaccountname"][0].ToString());
                                using (DirectoryEntry de = user.GetUnderlyingObject() as DirectoryEntry)
                                {
                                    de.RefreshCache(new string[] { "canonicalName" });
                                    string canonicalName = de.Properties["canonicalName"].Value as string;
                                    OutputBox.AppendText($"OU:\t\t{canonicalName}\n");
                                }
                            }
                        }
                        else if (PropertyList[i, 0] == "UR AD")
                        {
                            DirectoryEntry URentry = new DirectoryEntry("LDAP://ur.rochester.edu");
                            DirectorySearcher URsearcher = new DirectorySearcher(URentry);
                            URsearcher.Filter = "(&(objectClass=user)(uidNumber=" + UserResult.Properties["uidNumber"][0].ToString() + "))";
                            SearchResult URUserResult = URsearcher.FindOne();

                            if (URUserResult != null)
                            {
                                OutputBox.AppendText($"{PropertyList[i, 0]}:\t" + URUserResult.Properties["samaccountname"][0] + '\n');
                            }
                            else
                            {
                                OutputBox.AppendText($"{PropertyList[i, 0]}:\tNo UR AD account\n");
                            }
                        }
                        else
                                {
                            try
                            {
                                OutputBox.AppendText($"{PropertyList[i, 0]}\t" + UserResult.Properties[PropertyList[i, 1]][0] + '\n');
                            }
                            catch (Exception ex)
                            {
                                OutputBox.AppendText($"{PropertyList[i, 0]}\t[Not Listed]\n");
                            }
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
                                bool FullAccess = false;
                                bool SendAsAccess = false;
                                bool SendOnBehalfAccess = false;

                                foreach (var item in UserResult.Properties["memberof"])
                                {
                                    string test = item.ToString().Substring(3, item.ToString().IndexOf(",") - 3);
                                    string test2 = MailboxOwners[i];
                                    if (MailboxOwners[i].Contains(item.ToString().Substring(3, item.ToString().IndexOf(",") - 3)))
                                    {
                                        string[] LineArray = MailboxOwners[i].Split(',');

                                        for (int j = 0; j < LineArray.Count(); j++)
                                        {
                                            if (LineArray[j].Contains(item.ToString().Substring(3, item.ToString().IndexOf(",") - 3)))
                                            {
                                                if (j == 4)
                                                {
                                                    FullAccess = true;
                                                }
                                                else if (j == 5)
                                                {
                                                    SendAsAccess = true;
                                                }
                                                else if (j == 6)
                                                {
                                                    SendOnBehalfAccess = true;
                                                }
                                            }
                                        }
                                        HasMailboxOrDL = true;
                                    }
                                }

                                if (MailboxOwners[i].Contains(UserResult.Properties["samaccountname"][0].ToString()))
                                {
                                    string[] LineArray = MailboxOwners[i].Split(',');

                                    for (int j = 0; j < LineArray.Count(); j++)
                                    {
                                        if (LineArray[j].Contains(UserResult.Properties["samaccountname"][0].ToString()))
                                        {
                                            if (j == 4)
                                            {
                                                FullAccess = true;
                                            }
                                            else if (j == 5)
                                            {
                                                SendAsAccess = true;
                                            }
                                            else if (j == 6)
                                            {
                                                SendOnBehalfAccess = true;
                                            }
                                        }
                                    }
                                    HasMailboxOrDL = true;
                                }

                                if (FullAccess || SendAsAccess || SendOnBehalfAccess)
                                    OutputBox.AppendText("Mailbox:\t" + MailboxOwners[i].ToString().Substring(0, MailboxOwners[i].ToString().IndexOf(",")));
                                if (FullAccess)
                                    OutputBox.AppendText("\tFull Access");
                                if (SendAsAccess)
                                    OutputBox.AppendText("\tSend As Access");
                                if (SendOnBehalfAccess)
                                    OutputBox.AppendText("\tSend On Behalf Access");
                                if (FullAccess || SendAsAccess || SendOnBehalfAccess)
                                    OutputBox.AppendText("\n");

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
                                    OutputBox.AppendText("Mailbox:\t" + MailboxOwners[i].ToString().Substring(0, MailboxOwners[i].ToString().IndexOf(",")));
                                    string[] LineArray = MailboxOwners[i].Split(',');

                                    for (int j = 0; j < LineArray.Count(); j++)
                                    {
                                        if (LineArray[j].Contains(UserResult.Properties["mail"][0].ToString()))
                                        {
                                            if (j == 4)
                                            {
                                                OutputBox.AppendText("\tFull Access");
                                            }
                                            else if (j == 5)
                                            {
                                                OutputBox.AppendText("\tSend As Access");
                                            }
                                            else if (j == 6)
                                            {
                                                OutputBox.AppendText("\tSend On Behalf Access");
                                            }
                                        }
                                    }
                                    OutputBox.AppendText("\n");
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
                                    OutputBox.AppendText("Managed DL:\t" + DLOwners[i].ToString().Substring(0, DLOwners[i].ToString().IndexOf(",")) + "\n");
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
                    OutputBox.AppendText("\nShared/Home Drive Access List: \n");
                    try
                    {
                        OutputBox.AppendText("H" + UserResult.Properties["homedirectory"][0].ToString() + "\n");
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText($"{UserResult.Properties["name"][0]} has no H: Drive\n");
                    }
                    try
                    {
                        ResultPropertyValueCollection groups = UserResult.Properties["memberOf"];
                        using (var sr = new StreamReader("\\\\ADSDC01\\netlogon\\SIG\\logon.dmd"))
                        {
                            string[] ShareList = sr.ReadToEnd().Split("\n");
                            foreach (var group in groups)
                            {
                                for (int i = 0; i < ShareList.Length; i++)
                                {
                                    if (ShareList[i].Contains(group.ToString().Substring(3, group.ToString().IndexOf(",") - 3) + "|"))
                                    {
                                        OutputBox.AppendText(ShareList[i].Substring(ShareList[i].LastIndexOf('|') + 1, ShareList[i].Substring(ShareList[i].LastIndexOf('|')).Length - 2) + " | " + ShareList[i].Substring(ShareList[i].IndexOf("\\") + 1, ShareList[i].IndexOf('|') - 8) + '\n');
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText($"Unable to find Shared Drives for {UserName}");
                    }
                }
                else
                {
                    OutputBox.AppendText($"Unable to find username \"{UserName}\"");
                }
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ComputerInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (NameBox.Text != "")
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
                    string[,] PropertyList = { { "Domain Name", "DNSHostName" }, { "OS", "operatingsystem" }, { "OS Version", "operatingsystemversion" }, { "LAPS password", "ms-mcs-admpwd" } };

                    for (int i = 0; i < PropertyList.Length / 2; i++)
                    {
                        try
                        {
                            OutputBox.AppendText($"{PropertyList[i, 0]}: " + result.Properties[PropertyList[i, 1]][0] + '\n');
                        }
                        catch (Exception ex)
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
                        OutputBox.AppendText("OU: " + canonicalName + "\n");
                    }
                }
                else
                    OutputBox.AppendText($"{ComputerName} not found");
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void GroupInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserTextbox.Text != "")
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
                        catch (Exception ex)
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
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                "\nThe Master Seach Button Queries URMC and UR servers and the print Q report for any items that match" +
                "\nPlease reach out to Joseph Patrick on Teams or joseph_patrick@urmc.rochester.edu / jpatrick5402@gmail.com for any questions comments or concerns";
            OutputBox.AppendText(helpInfo);
            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void DarkButton_Click(object sender, RoutedEventArgs e)
        {
            if (OutputBox.Background != Brushes.Black)
            {
                MainWindow1.Background = new LinearGradientBrush(Colors.Navy, Colors.Black, 90.00);
                OutputBox.Background = Brushes.Black;
                OutputBox.Foreground = Brushes.White;
                PCNameBox.Foreground = Brushes.White;
                ADNameBox.Foreground = Brushes.White;
                NameBox.Background = Brushes.Black;
                NameBox.Foreground = Brushes.White;
                UserTextbox.Background = Brushes.Black;
                UserTextbox.Foreground = Brushes.White;

                RestartButton.Background = Brushes.Black;
                RestartButton.Foreground = Brushes.White;
                PingButton.Background = Brushes.Black;
                PingButton.Foreground = Brushes.White;
                SysinfoButton.Background = Brushes.Black;
                SysinfoButton.Foreground = Brushes.White;
                NslookupButton.Background = Brushes.Black;
                NslookupButton.Foreground = Brushes.White;
                RemoteButton.Background = Brushes.Black;
                RemoteButton.Foreground = Brushes.White;
                ComputerInfoButton.Background = Brushes.Black;
                ComputerInfoButton.Foreground = Brushes.White;
                UserInfoButton.Background = Brushes.Black;
                UserInfoButton.Foreground = Brushes.White;
                MemberOfButton.Background = Brushes.Black;
                MemberOfButton.Foreground = Brushes.White;
                GroupInfoButton.Background = Brushes.Black;
                GroupInfoButton.Foreground = Brushes.White;
                GroupMembersButton.Background = Brushes.Black;
                GroupMembersButton.Foreground = Brushes.White;
                AddgroupButton.Background = Brushes.Black;
                AddgroupButton.Foreground = Brushes.White;
                HelpButton.Background = Brushes.Black;
                HelpButton.Foreground = Brushes.White;
                DarkButton.Background = Brushes.Black;
                DarkButton.Foreground = Brushes.White;
                ExportButton.Background = Brushes.Black;
                ExportButton.Foreground = Brushes.White;
                LockoutToolButton.Background = Brushes.Black;
                LockoutToolButton.Foreground = Brushes.White;
                ToolTab1.Background = Brushes.Black;
                ToolTab1.Foreground = Brushes.DarkGray;
                ToolTab2.Background = Brushes.Black;
                ToolTab2.Foreground = Brushes.DarkGray;
                ToolTab3.Background = Brushes.Black;
                ToolTab3.Foreground = Brushes.DarkGray;
                ToolTab5.Background = Brushes.Black;
                ToolTab5.Foreground = Brushes.DarkGray;
                ToolBox1.Background = Brushes.Black;
                ToolBox2.Background = Brushes.Black;
                ToolBox3.Background = Brushes.Black;
                ToolBox5.Background = Brushes.Black;
                URMCDomainCB.Foreground = Brushes.White;
                URDomainCB.Foreground = Brushes.White;
                PrintersCB.Foreground = Brushes.White;
                SharedDriveCB.Foreground = Brushes.White;
                MasterSearchBox.Background = Brushes.Black;
                MasterSearchBox.Foreground = Brushes.White;
                Settings.Default.DarkMode = true;
                Settings.Default.Save();
            }
            else
            {
                MainWindow1.Background = new LinearGradientBrush(Colors.White, Colors.Yellow, 90.00);
                OutputBox.Background = Brushes.White;
                OutputBox.Foreground = Brushes.Black;
                PCNameBox.Foreground = Brushes.Black;
                ADNameBox.Foreground = Brushes.Black;
                NameBox.Background = Brushes.White;
                NameBox.Foreground = Brushes.Black;
                UserTextbox.Background = Brushes.White;
                UserTextbox.Foreground = Brushes.Black;

                RestartButton.Background = Brushes.White;
                RestartButton.Foreground = Brushes.Black;
                PingButton.Background = Brushes.White;
                PingButton.Foreground = Brushes.Black;
                SysinfoButton.Background = Brushes.White;
                SysinfoButton.Foreground = Brushes.Black;
                NslookupButton.Background = Brushes.White;
                NslookupButton.Foreground = Brushes.Black;
                RemoteButton.Background = Brushes.White;
                RemoteButton.Foreground = Brushes.Black;
                ComputerInfoButton.Background = Brushes.White;
                ComputerInfoButton.Foreground = Brushes.Black;
                UserInfoButton.Background = Brushes.White;
                UserInfoButton.Foreground = Brushes.Black;
                MemberOfButton.Background = Brushes.White;
                MemberOfButton.Foreground = Brushes.Black;
                GroupInfoButton.Background = Brushes.White;
                GroupInfoButton.Foreground = Brushes.Black;
                GroupMembersButton.Background = Brushes.White;
                GroupMembersButton.Foreground = Brushes.Black;
                AddgroupButton.Background = Brushes.White;
                AddgroupButton.Foreground = Brushes.Black;
                HelpButton.Background = Brushes.White;
                HelpButton.Foreground = Brushes.Black;
                DarkButton.Background = Brushes.White;
                DarkButton.Foreground = Brushes.Black;
                ExportButton.Background = Brushes.White;
                ExportButton.Foreground = Brushes.Black;
                LockoutToolButton.Background = Brushes.White;
                LockoutToolButton.Foreground = Brushes.Black;
                ToolTab1.Background = Brushes.White;
                ToolTab1.Foreground = Brushes.Black;
                ToolTab2.Background = Brushes.White;
                ToolTab2.Foreground = Brushes.Black;
                ToolTab3.Background = Brushes.White;
                ToolTab3.Foreground = Brushes.Black;
                ToolTab5.Background = Brushes.White;
                ToolTab5.Foreground = Brushes.Black;
                ToolBox1.Background = Brushes.White;
                ToolBox2.Background = Brushes.White;
                ToolBox3.Background = Brushes.White;
                ToolBox5.Background = Brushes.White;
                URMCDomainCB.Foreground = Brushes.Black;
                URDomainCB.Foreground = Brushes.Black;
                PrintersCB.Foreground = Brushes.Black;
                SharedDriveCB.Foreground = Brushes.Black;
                MasterSearchBox.Background = Brushes.White;
                MasterSearchBox.Foreground = Brushes.Black;
                Settings.Default.DarkMode = false;
                Settings.Default.Save();
            }
        }
        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            TextRange textRange = new TextRange(OutputBox.Document.ContentStart, OutputBox.Document.ContentEnd);
            string[] OutputArray = textRange.Text.Split('\n');
            Directory.CreateDirectory($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\HDTEXPORT\\");
            string path = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\HDTEXPORT\\HDT_Export_{DateTime.Now.ToString("M-d-yyyy HH-mm-ss")}.csv";

            using (StreamWriter sw = new StreamWriter(path))
            {
                foreach (var line in OutputArray)
                {
                    sw.WriteLine(line);
                }
            }

            OutputBox.AppendText($"Export completed to {path}");

            OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        private void LockoutToolButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserTextbox.Text != "")
            {
                var UserName = UserTextbox.Text.Trim();

                MessageBoxResult result = System.Windows.MessageBox.Show("This may take approximately 1 minute, do you wish to continue?",
                                          "Confirmation",
                                          MessageBoxButton.YesNo,
                                          MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {

                    LoadingWindow Window = ShowLoadingWindow();
                    OutputBox.AppendText("Searching Bad password attempts for " + UserName + "...\n\n");
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
                        try
                        {
                            string[] DCs = { "ADPDC01", "ADPDC02", "ADPDC03", "ADPDC04", "ADPDC05", "ADSDC01", "ADSDC02", "ADSDC03", "ADSDC04", "ADSDC05" };
                            OutputBox.AppendText("DC\t\tCount\tTime\t\t\t\tLast Set\n");
                            foreach (string DC in DCs)
                            {
                                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, DC))
                                {
                                    UserPrincipal auser = UserPrincipal.FindByIdentity(context, UserResult.Properties["samaccountname"][0].ToString());
                                    var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                    DateTime LastBad = TimeZoneInfo.ConvertTime((DateTime)auser.LastBadPasswordAttempt, timeZone);
                                    DateTime LastSet = TimeZoneInfo.ConvertTime((DateTime)auser.LastPasswordSet, timeZone);
                                    if (!timeZone.IsDaylightSavingTime(LastBad))
                                    {
                                        LastBad = LastBad.AddHours(1);
                                    }
                                    if (!timeZone.IsDaylightSavingTime(LastSet))
                                    {
                                        LastSet = LastSet.AddHours(1);
                                    }
                                    OutputBox.AppendText(DC + "\t" + auser.BadLogonCount + "\t" + LastBad.ToString() + "\t\t" + LastSet.ToString() + "\n");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            OutputBox.AppendText($"An error occurred when trying to get password information for {UserName}");
                        }
                    }
                    else
                    {
                        OutputBox.AppendText($"Unable to find username \"{UserName}\"");
                    }
                    CloseLoadingWindow(Window);
                }
                else
                {
                    OutputBox.AppendText("Action Cancelled\n");
                }
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToEnd();
            }
            else
            {
                System.Windows.MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void FuzzySearch(object sender, System.Windows.Input.KeyEventArgs e)
        {
            var SearchObject = MasterSearchBox.Text.Trim();
            if (SearchObject.Length > 3)
            {
                OutputBox.Document.Blocks.Clear();
                SearchObject = MasterSearchBox.Text.Trim();
                OutputBox.AppendText("Searching for \"" + SearchObject + "*\"...\n\n\n");
                if (URMCDomainCB.IsChecked == true)
                {
                    try
                    {
                        // Search under URMC umbrella
                        DirectoryEntry entry = new DirectoryEntry("LDAP://urmc-sh.rochester.edu");
                        DirectorySearcher searcher = new DirectorySearcher(entry);

                        bool ResultIsFound = false;

                        searcher.Filter = $"(anr={SearchObject})";
                        SearchResultCollection UserResult = searcher.FindAll();
                        int count = 0;
                        foreach (SearchResult result in UserResult)
                        {
                            OutputBox.AppendText("URMC:   ");
                            OutputBox.AppendText(result.Properties["cn"][0].ToString().PadRight(35));
                            try { OutputBox.AppendText(result.Properties["uidNumber"][^1].ToString().PadRight(15)); }
                            catch { OutputBox.AppendText("[No URID]".PadRight(15)); }
                            OutputBox.AppendText(result.Properties["objectclass"][^1].ToString().PadRight(15));
                            try { OutputBox.AppendText(result.Properties["description"][0] + "\n"); }
                            catch { OutputBox.AppendText("[No Description]\n"); }
                            ResultIsFound = true;
                            if (count == 10)
                                break;
                            else count++;
                        }

                        if (!ResultIsFound) OutputBox.AppendText("URMC:\tNo object found\n\n");
                        else OutputBox.AppendText("\n");
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText($"An error has occurred, try refining your search or adjusting your search criteria: {ex.Message}\n\n");
                    }
                }
                if (URDomainCB.IsChecked == true)
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://ur.rochester.edu");
                        DirectorySearcher searcher = new DirectorySearcher(entry);

                        bool ResultIsFound = false;

                        searcher.Filter = $"(anr={SearchObject})";
                        SearchResultCollection UserResult = searcher.FindAll();
                        int count = 0;
                        foreach (SearchResult result in UserResult)
                        {
                            OutputBox.AppendText("UR:     ");
                            OutputBox.AppendText(result.Properties["cn"][0].ToString().PadRight(35));
                            try { OutputBox.AppendText(result.Properties["uidNumber"][^1].ToString().PadRight(15)); }
                            catch { OutputBox.AppendText("[No URID]".PadRight(15)); }
                            OutputBox.AppendText(result.Properties["objectclass"][^1].ToString().PadRight(15));
                            try { OutputBox.AppendText(result.Properties["description"][0] + "\n"); }
                            catch { OutputBox.AppendText("[No Description]\n"); }
                            ResultIsFound = true;
                            if (count == 10)
                                break;
                            else count++;
                        }

                        if (!ResultIsFound) OutputBox.AppendText("UR:\tNo object found\n\n");
                        else OutputBox.AppendText("\n");
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText($"An error has occurred, try refining your search or adjusting your search criteria: {ex.Message}\n\n");
                    }
                }
                if (SharedDriveCB.IsChecked == true)
                {
                    try
                    {
                        using (var sr = new StreamReader("\\\\ADSDC01\\netlogon\\SIG\\logon.dmd"))
                        {
                            string[] ShareList = sr.ReadToEnd().Split("\n");
                            bool ItemFound = false;
                            int count = 0;
                            for (int i = 0; i < ShareList.Length; i++)
                            {
                                if (ShareList[i].Contains(SearchObject))
                                {
                                    OutputBox.AppendText(ShareList[i].Substring(ShareList[i].LastIndexOf('|') + 1, ShareList[i].Substring(ShareList[i].LastIndexOf('|')).Length - 2) + " | " + ShareList[i].Substring(ShareList[i].IndexOf("\\") + 1, ShareList[i].IndexOf('|') - 8) + '\n');
                                    ItemFound = true;
                                    if (count == 10)
                                        break;
                                    else count++;
                                }
                            }
                            if (!ItemFound)
                                OutputBox.AppendText("No Shares found with criteria\n\n");
                            else
                                OutputBox.AppendText("\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText($"Unable to find Shared Drives: {ex.Message}\n\n");
                    }
                }
                if (PrintersCB.IsChecked == true)
                {
                    try
                    {
                        // This section is used to see if there are any printers that match the search criteria as well
                        string url = "https://apps.mc.rochester.edu/ISD/SIG/PrintQueues/PrintQReport.csv";
                        using (HttpClient client = new HttpClient())
                        {
                            bool ResultIsFound = false;
                            HttpResponseMessage response = await client.GetAsync(url);
                            if (response.IsSuccessStatusCode)
                            {
                                string csvContent = await response.Content.ReadAsStringAsync();
                                string[] StringArray = csvContent.Split('\n');
                                int count = 0;
                                OutputBox.AppendText("Ignore if Duplicate\n");
                                for (int i = 0; i < StringArray.Length; i++)
                                {
                                    if (StringArray[i].Contains($"{SearchObject}"))
                                    {
                                        OutputBox.AppendText(StringArray[i].Replace("\"", "").Replace(",", ", "));
                                        ResultIsFound = true;
                                        if (count == 10)
                                            break;
                                        else count++;           
                                    }
                                }
                                if (!ResultIsFound)
                                    OutputBox.AppendText("No printers found with criteria\n\n");
                                else
                                    OutputBox.AppendText("\n");
                            }
                            else
                            {
                                OutputBox.AppendText($"Failed to download CSV. Status code: {response.StatusCode}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OutputBox.AppendText($"An error has occurred {ex.Message}\n\n");
                    }
                }
                OutputBox.AppendText("\n-------------------------------------------------------------------------------------------------------------------------\n");
                OutputBox.ScrollToHome();
            }
        }
    }
}