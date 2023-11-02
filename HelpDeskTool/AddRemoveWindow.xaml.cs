using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace DTTool
{
    /// <summary>
    /// Interaction logic for AddRemoveWindow.xaml
    /// </summary>
    public partial class AddRemoveWindow : Window
    {
        public AddRemoveWindow()
        {
            InitializeComponent();
            //ARoutputBox.Document.Blocks.Clear();
        }
        // Remove Button
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ARoutputbox.Document.Blocks.Clear();
            var user = ARusernameBox.Text;
            var group = ARgroupBox.Text;
            if (user != "" && group != "")
            {
                ARusernameBox.Clear();
                ARgroupBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                // Remove-ADGroupMember
                // Remove-ADGroupMember -Identity GROUP -Members USERNAME
                command.StartInfo.Arguments = "/C powershell Remove-ADGroupMember \'" + group + "\' \'" + user + "\' -Confirm:$false";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                ARoutputbox.AppendText(console_output);
                if (command.ExitCode == 0)
                {
                    ARoutputbox.AppendText(user + " Removed from " + group + "\n");
                    ARoutputbox.AppendText("May take up to 30 seconds for changes to reflect in AD");
                }
                else
                {
                    ARoutputbox.AppendText("\nAction Failed\n");
                }
            }
            else ARoutputbox.AppendText("No Input Detected\n");
            ARoutputbox.ScrollToEnd();
        }
        // Add Button
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ARoutputbox.Document.Blocks.Clear();
            var user = ARusernameBox.Text;
            var group = ARgroupBox.Text;
            if (user != "" && group != "")
            {
                ARusernameBox.Clear();
                ARgroupBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                // Add-ADGroupMember
                // Add-ADGroupMember -Identity GROUP -Members USERNAME
                command.StartInfo.Arguments = "/C powershell Add-ADGroupMember \'" + group + "\' \'" + user + "\'";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                ARoutputbox.AppendText(console_output);
                if (command.ExitCode == 0)
                {
                    ARoutputbox.AppendText(user + " Added to " + group + "\n");
                    ARoutputbox.AppendText("May take up to 30 seconds for changes to reflect in AD");
                }
                else
                {
                    ARoutputbox.AppendText("\nAction Failed\n");
                }
            }
            else ARoutputbox.AppendText("No Input Detected\n");
            ARoutputbox.ScrollToEnd();
        }
        // Add ISDG_CTX_eRecord2
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            ARoutputbox.Document.Blocks.Clear();
            var user = ARusernameBox.Text;
            if (user != "")
            {
                ARusernameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                // Add-ADGroupMember
                // Add-ADGroupMember -Identity GROUP -Members USERNAME
                command.StartInfo.Arguments = "/C powershell Add-ADGroupMember \'ISDG_CTX_eRecord2\' \'" + user + "\'";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                ARoutputbox.AppendText("Exit Code: " + command.ExitCode.ToString() + "\n\n");
                if (command.ExitCode == 0)
                {
                    ARoutputbox.AppendText(user + " Added to ISDG_CTX_eRecord2\n");
                    ARoutputbox.AppendText("May take up to 30 seconds for changes to reflect in AD");
                }
                else
                {
                    ARoutputbox.AppendText("\nAction Failed\n");
                }
            }
            else ARoutputbox.AppendText("No Username Detected\n");
            ARoutputbox.ScrollToEnd();
        }
        // Add ISDU_CitrixAccessGateway
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            ARoutputbox.Document.Blocks.Clear();
            var user = ARusernameBox.Text;
            if (user != "")
            {
                ARusernameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                // Add-ADGroupMember
                // Add-ADGroupMember -Identity GROUP -Members USERNAME
                command.StartInfo.Arguments = "/C powershell Add-ADGroupMember \'ISDU_CitrixAccessGateway\' \'" + user + "\'";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                ARoutputbox.AppendText("Exit Code: " + command.ExitCode.ToString() + "\n\n");
                if (command.ExitCode == 0)
                {
                    ARoutputbox.AppendText(user + " Added to ISDU_CitrixAccessGateway\n");
                    ARoutputbox.AppendText("May take up to 30 seconds for changes to reflect in AD");
                }
                else
                {
                    ARoutputbox.AppendText("\nAction Failed\n");
                }
            }
            else ARoutputbox.AppendText("No Username Detected\n");
            ARoutputbox.ScrollToEnd();
        }
        // Cisco AnyConnect Access
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            ARoutputbox.Document.Blocks.Clear();
            var user = ARusernameBox.Text;
            if (user != "")
            {
                ARusernameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                // Add-ADGroupMember
                // Add-ADGroupMember -Identity GROUP -Members USERNAME
                command.StartInfo.Arguments = "/C powershell Add-ADGroupMember \'ISDG_VPN_FullAccess\' \'" + user + "\'";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                ARoutputbox.AppendText("Exit Code: " + command.ExitCode.ToString() + "\n\n");
                if (command.ExitCode == 0)
                {
                    ARoutputbox.AppendText(user + " Added to ISDG_VPN_FullAccess\n");
                    ARoutputbox.AppendText("May take up to 30 seconds for changes to reflect in AD");
                }
                else
                {
                    ARoutputbox.AppendText("\nAction Failed\n");
                }
            }
            else ARoutputbox.AppendText("No Username Detected\n");
            ARoutputbox.ScrollToEnd();
        }
        // GP Access
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            ARoutputbox.Document.Blocks.Clear();
            var user = ARusernameBox.Text;
            if (user != "")
            {
                ARusernameBox.Clear();
                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "cmd";
                // Add-ADGroupMember
                // Add-ADGroupMember -Identity GROUP -Members USERNAME
                command.StartInfo.Arguments = "/C powershell Add-ADGroupMember \'ISDU_VPN_GP_FullAccess\' \'" + user + "\'";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                ARoutputbox.AppendText("Exit Code: " + command.ExitCode.ToString() + "\n\n");
                if (command.ExitCode == 0)
                {
                    ARoutputbox.AppendText(user + " Added to ISDU_VPN_GP_FullAccess\n");
                    ARoutputbox.AppendText("May take up to 30 seconds for changes to reflect in AD");
                }
                else
                {
                    ARoutputbox.AppendText("\nAction Failed\n");
                }
            }
            else ARoutputbox.AppendText("No Username Detected\n");
            ARoutputbox.ScrollToEnd();
        }
    }
}
