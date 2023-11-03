using Microsoft.VisualBasic.ApplicationServices;
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
                return false;
            }
            else
            {
                return true;
            }
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

        public void AddOrRemove(string AddorRemove)
        {
            if (RTPHasText(ARusernameBox) && RTPHasText(ARgroupBox))
            {
                string usertext = StringFromRichTextBox(ARusernameBox);
                string grouptext = StringFromRichTextBox(ARgroupBox);
                string errString = "";

                var myUserList = usertext.Split("\r\n");
                var myGroupList = grouptext.Split("\r\n");

                aProgressBar.Maximum = myUserList.Length + myGroupList.Length - 2;

                foreach (var group in myGroupList)
                {
                    aProgressBar.Value++;

                    foreach (var user in myUserList)
                    {
                        if (user == "" || group == "")
                            break;

                        aProgressBar.Value++;

                        group.Trim();
                        user.Trim();

                        System.Diagnostics.Process command = new System.Diagnostics.Process();
                        command.StartInfo.CreateNoWindow = true;
                        command.StartInfo.FileName = "cmd";
                        if (AddorRemove == "Remove")
                        {
                            command.StartInfo.Arguments = "/C powershell Remove-ADGroupMember \'" + group + "\' \'" + user + "\' -Confirm:$false";
                        }else if(AddorRemove == "Add")
                        {
                            command.StartInfo.Arguments = "/C powershell Add-ADGroupMember \'" + group + "\' \'" + user + "\'";
                        }
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
                    MessageBox.Show("All users processed successfully\nMay take up to 30 seconds to reflect in AD", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    ARusernameBox.Document.Blocks.Clear();
                    ARgroupBox.Document.Blocks.Clear();
                    aProgressBar.Value = 0;
                }
                else
                {
                    MessageBox.Show(errString, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Input Missing", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Remove Button
        private void RemoveUGButton_Click(object sender, RoutedEventArgs e)
        {
            AddOrRemove("Remove");
        }

        // Add Button
        private void AddUGButton_Click(object sender, RoutedEventArgs e)
        {
            AddOrRemove("Add");
        }

        private void AddGP_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AddCisco_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AddMyApps_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AddeRecord_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
