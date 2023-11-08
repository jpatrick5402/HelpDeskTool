﻿using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
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
    /// There are two main functions, one is AddOrRemoveBulk, meaning add every user in the user box to every group in the group box.
    /// There is AddGroup which adds each user to the group specified
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
        public string StringFromRTB(RichTextBox rtb)
        {
            TextRange textRange = new TextRange(
                rtb.Document.ContentStart,
                rtb.Document.ContentEnd
            );
            return textRange.Text;
        }

        public void AddOrRemoveBulk(string AddorRemove)
        {
            if (RTPHasText(ARusernameBox) && RTPHasText(ARgroupBox))
            {
                string usertext = StringFromRTB(ARusernameBox);
                string grouptext = StringFromRTB(ARgroupBox);
                string errString = "";
                string userStringFormatted = "";

                string [] myUserList = usertext.Split("\r\n");
                string [] myGroupList = grouptext.Split("\r\n");

                foreach(string user in myUserList)
                {
                    string userClean = user.Trim();
                    userStringFormatted += userClean + ",";
                }

                if (AddorRemove == "Remove")
                {
                    foreach (var group in myGroupList)
                    {
                        if (group == "")
                            break;

                        string groupClean = group.Trim();

                        System.Diagnostics.Process command = new System.Diagnostics.Process();
                        command.StartInfo.CreateNoWindow = true;
                        command.StartInfo.FileName = "powershell";
                        command.StartInfo.Arguments = "Remove-ADGroupMember \'" + groupClean + "\' \'" + userStringFormatted + "\' -Confirm:$false";
                        command.StartInfo.RedirectStandardOutput = true;
                        command.Start();
                        var console_output = command.StandardOutput.ReadToEnd();
                        if (command.ExitCode != 0)
                        {
                            using (var reader = new StringReader(console_output))
                            {
                                string first = reader.ReadLine();
                                errString += first + "\n";
                            }
                        }

                    }
                }
                else if (AddorRemove == "Add")
                {
                    foreach (var group in myGroupList)
                    {
                        if (group == "")
                            break;

                        string groupClean = group.Trim();

                        System.Diagnostics.Process command = new System.Diagnostics.Process();
                        command.StartInfo.CreateNoWindow = true;
                        command.StartInfo.FileName = "powershell";
                        command.StartInfo.Arguments = "Add-ADGroupMember \'" + groupClean + "\' \'" + userStringFormatted + "\'";
                        command.StartInfo.RedirectStandardOutput = true;
                        command.Start();
                        var console_output = command.StandardOutput.ReadToEnd();
                        if (command.ExitCode != 0)
                        {
                            using (var reader = new StringReader(console_output))
                            {
                                string first = reader.ReadLine();
                                errString += first + "\n";
                            }
                        }

                    }
                }
                if (errString == "")
                {
                    MessageBox.Show("All users processed successfully\nMay take up to 30 seconds to reflect in AD", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    ARusernameBox.Document.Blocks.Clear();
                    ARgroupBox.Document.Blocks.Clear();
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

        // add group to speed up process to and a singular group to users
        private void AddGroup(string group)
        {
            if (RTPHasText(ARusernameBox))
            {
                string usertext = StringFromRTB(ARusernameBox);
                string errString = "";
                string [] myUserList = usertext.Split("\r\n");
                string userStringFormatted = "";

                foreach (string user in myUserList)
                {
                    string userClean = user.Trim();
                    userStringFormatted += userClean + ",";
                }

                System.Diagnostics.Process command = new System.Diagnostics.Process();
                command.StartInfo.CreateNoWindow = true;
                command.StartInfo.FileName = "powershell";
                command.StartInfo.Arguments = "Add-ADGroupMember \'" + group + "\' \'" + userStringFormatted + "\'";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                var console_output = command.StandardOutput.ReadToEnd();
                if (command.ExitCode != 0)
                {
                    using (var reader = new StringReader(console_output))
                    {
                        string first = reader.ReadLine();
                        errString += first + "\n";
                    }
                }
                if (errString == "")
                {
                    MessageBox.Show("All users processed successfully\nMay take up to 30 seconds to reflect in AD", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    ARusernameBox.Document.Blocks.Clear();
                }
                else
                {
                    MessageBox.Show(errString, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("User Input Missing", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Remove Button
        private void RemoveUGButton_Click(object sender, RoutedEventArgs e)
        {
            AddOrRemoveBulk("Remove");
        }

        // Add Button
        private void AddUGButton_Click(object sender, RoutedEventArgs e)
        {
            AddOrRemoveBulk("Add");
        }

        private void AddGP_Click(object sender, RoutedEventArgs e)
        {
            AddGroup("ISDU_VPN_GP_FullAccess");
        }

        private void AddCisco_Click(object sender, RoutedEventArgs e)
        {
            AddGroup("ISDG_VPN_FullAccess");
        }

        private void AddMyApps_Click(object sender, RoutedEventArgs e)
        {
            AddGroup("ISDU_CitrixAccessGateway");
        }

        private void AddeRecord_Click(object sender, RoutedEventArgs e)
        {
            AddGroup("ISDG_CTX_eRecord2");
        }
    }
}
