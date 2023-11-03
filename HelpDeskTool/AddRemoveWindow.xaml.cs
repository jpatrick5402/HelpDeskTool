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

        public bool RTPHasText(RichTextBox rtb)
        {
            if (rtb.Document.Blocks.Count == 0) return true;
            TextPointer startPointer = rtb.Document.ContentStart.GetNextInsertionPosition(LogicalDirection.Forward);
            TextPointer endPointer = rtb.Document.ContentEnd.GetNextInsertionPosition(LogicalDirection.Backward);
            if (startPointer.CompareTo(endPointer) == 0)
            {
                MessageBox.Show("No text in " + rtb.Name, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            return false;
        }
        public string StringFromRichTextBox(RichTextBox rtb)
        {
            TextRange textRange = new TextRange(
                // TextPointer to the start of content in the RichTextBox.
                rtb.Document.ContentStart,
                // TextPointer to the end of content in the RichTextBox.
                rtb.Document.ContentEnd
            );

            // The Text property on a TextRange object returns a string
            // representing the plain text content of the TextRange.
            return textRange.Text;
        }

        // Remove Button
        private void RemoveUGButton_Click(object sender, RoutedEventArgs e)
        {
            RTPHasText(ARusernameBox);
            RTPHasText(ARgroupBox);

            string usertext = StringFromRichTextBox(ARusernameBox);
            string grouptext = StringFromRichTextBox(ARgroupBox);
            string errString = "";

            var myUserList = usertext.Split("\r\n");
            var myGroupList = usertext.Split("\r\n");

            foreach (var group in myGroupList)
            {
                foreach (var user in myUserList)
                {
                    group.Trim();
                    user.Trim();
                    System.Diagnostics.Process command = new System.Diagnostics.Process();
                    command.StartInfo.CreateNoWindow = true;
                    command.StartInfo.FileName = "cmd";
                    command.StartInfo.Arguments = "/C powershell Remove-ADGroupMember \'" + group + "\' \'" + user + "\' -Confirm:$false";
                    command.StartInfo.RedirectStandardOutput = true;
                    command.Start();
                    var console_output = command.StandardOutput.ReadToEnd();
                    if (command.ExitCode != 0)
                    {
                        errString += console_output + "\n";
                    }
                }
            }
            if (errString == "")
            {
                MessageBox.Show("All users deleted", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(errString, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Add Button
        private void AddUGButton_Click(object sender, RoutedEventArgs e)
        {

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

        }

        private void RemoveGUButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
