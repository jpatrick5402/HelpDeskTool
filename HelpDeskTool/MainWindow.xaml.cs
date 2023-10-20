/// Program HDTool written by Joseph Patrick for use in the URMC ISD Help Desk environment
/// Created 9/25/2023

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
            OutputBox.AppendText("\nHelp Desk Tool\n");
            OutputBox.AppendText("\nAwaiting Commands\n");
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
        }

        public Boolean IsTextInNameBox()
        {
            if (NameBox.Text != "")
                return true;
            OutputBox.AppendText("No PC Name/IP Detected");
            return false;
        }
        public Boolean IsTextInUserBox()
        {
            if (UserTextbox.Text != "")
                return true;
            OutputBox.AppendText("No AD Name Detected\n");
            return false;
        }
        //Restart Button
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text;
                Clipboard.SetText(PCName);
                NameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                command.StartInfo.Arguments = "/C shutdown -r -t 2 -m " + PCName;
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText("Restart Initiated");
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Ping Button
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text;
                OutputBox.AppendText("Pinging " + PCName + "\n\n");

                Clipboard.SetText(PCName);
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
        //nslookup Button
        private void NslookupButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text;
                OutputBox.AppendText("Looking for " + PCName + "\n\n");

                Clipboard.SetText(PCName);
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
        //Sys Info Button
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text;
                OutputBox.AppendText("Gathering info on " + PCName + "\n\n");

                Clipboard.SetText(PCName);
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
        //Remote in Button
        private void RemoteButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text;
                Clipboard.SetText(PCName);
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
        //AD info button
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var Username = UserTextbox.Text;
                OutputBox.AppendText("Gathering info for " + Username + "\n\n");

                Clipboard.SetText(Username);
                UserTextbox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "powershell";
                command.StartInfo.Arguments = "Get-ADUser " + Username + " -Properties *";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Member of Button
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var Username = UserTextbox.Text;
                OutputBox.AppendText("Gathering groups for " + Username + "\n\n");

                Clipboard.SetText(Username);
                UserTextbox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "powershell";
                command.StartInfo.Arguments = "Get-ADPrincipalGroupMembership " + Username + " | select name";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd().Replace(" ", ""));
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Get Serial Button
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                var PCName = NameBox.Text;
                OutputBox.AppendText("Gathering Serial Number for " + PCName + "\n\n");

                Clipboard.SetText(PCName);
                NameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                command.StartInfo.Arguments = "/C powershell Get-WMIObject Win32_Bios -ComputerName " + PCName + " | find \"SerialNumber\"";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        // Members Button
        // This button uses the powershell command "Get-ADGroupMember to list the usernames and names of a group, this allows for easy copy/paste into AD or HDAMU
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var Username = UserTextbox.Text;
                OutputBox.AppendText("Gathering Group Members for " + Username + "\n\n");

                // Gather usernames
                Clipboard.SetText(Username);
                UserTextbox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "powershell";
                command.StartInfo.Arguments = "Get-ADGroupMember -Identity \'" + Username + "\' | select SamAccountName";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd().Replace(" ", ""));
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Add group button
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            AddRemoveWindow win = new AddRemoveWindow();
            win.Show();
        }
        // Pass Last Set Button
        // Runs powershell commands to get AD info which includes the user's Password set date and when it last failed
        // There is a current issue determining the accuracy of the Last Bad Password fucntion because there are more that one DC's to use
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                var Username = UserTextbox.Text;
                OutputBox.AppendText("Gathering Password information for " + Username + "\n\n");

                Clipboard.SetText(Username);
                UserTextbox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "powershell";
                command.StartInfo.Arguments = "Get-ADUser -Identity " + Username + " -Properties PasswordLastSet, LastBadPasswordAttempt | select PasswordLastSet, LastBadPasswordAttempt | format-list";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());

                OutputBox.AppendText("\nLastBadPasswordAttempt NOT ALWAYS ACCURATE\n");
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        // Clear Button
        // Completely deletes all text from the main text box
        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            OutputBox.Document.Blocks.Clear();
        }
    }
}
