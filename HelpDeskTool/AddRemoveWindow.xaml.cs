using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection.Metadata;
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
        }

        public bool isRTBEmpty(RichTextBox rtb)
        {
            if (rtb.Document.Blocks.Count == 0) return true;
            TextPointer startPointer = rtb.Document.ContentStart.GetNextInsertionPosition(LogicalDirection.Forward);
            TextPointer endPointer = rtb.Document.ContentEnd.GetNextInsertionPosition(LogicalDirection.Backward);
            if (startPointer.CompareTo(endPointer) == 0)
            {
                return true;
            }
            else
            {
                MessageBox.Show("No text in " + rtb.Name, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            
        }

        // Remove Button
        private void RemoveUGButton_Click(object sender, RoutedEventArgs e)
        {
            isRTBEmpty(ARusernameBox);
            isRTBEmpty(ARgroupBox);


        /* Prevously used code
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
        */
        }
        // Add Button
        private void AddUGButton_Click(object sender, RoutedEventArgs e)
        {
            isRTBEmpty(ARusernameBox);
            isRTBEmpty(ARgroupBox);

            /* Prevously used code
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
            */
        }
        

        private void AddGUButton_Click(object sender, RoutedEventArgs e)
        {
            isRTBEmpty(ARusernameBox);
            isRTBEmpty(ARgroupBox);
        }

        private void RemoveGUButton_Click(object sender, RoutedEventArgs e)
        {
            isRTBEmpty(ARusernameBox);
            isRTBEmpty(ARgroupBox);
        }
    }
}
