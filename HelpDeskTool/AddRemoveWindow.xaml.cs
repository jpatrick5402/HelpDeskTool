using HelpDeskTool;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

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

        public static bool RTPHasText(RichTextBox rtb)
        {
            if (rtb.Document.Blocks.Count == 0)
            {
                return true;
            }
            TextPointer startPointer = rtb.Document.ContentStart.GetNextInsertionPosition(LogicalDirection.Forward);
            TextPointer endPointer = rtb.Document.ContentEnd.GetNextInsertionPosition(LogicalDirection.Backward);
            if (startPointer.CompareTo(endPointer) == 0)
            {
                return false;
            }
            return true;
        }

        public static string StringFromRTB(RichTextBox rtb)
        {
            TextRange textRange = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            return textRange.Text;
        }

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

        public void AddOrRemoveBulk(string[] users, string[] groups, string AddorRemove)
        {
            if (RTPHasText(ARusernameBox) && RTPHasText(ARgroupBox))
            {
                LoadingWindow Window = ShowLoadingWindow();
                string ErrorList = "";

                using PrincipalContext context = new PrincipalContext(ContextType.Domain, "urmc-sh.rochester.edu");
                foreach (string group in groups)
                {
                    if (group != "")
                    {
                        GroupPrincipal agroup = GroupPrincipal.FindByIdentity(context, group.Trim());
                        if (agroup == null)
                        {
                            MessageBox.Show($"Group not found: {group}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        else
                        {
                            foreach (string user in users)
                            {
                                UserPrincipal auser = UserPrincipal.FindByIdentity(context, user.Trim());
                                if (auser == null)
                                {
                                    MessageBox.Show($"User not found: {user}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else
                                {
                                    try
                                    {
                                        if (AddorRemove.ToLower() == "add")
                                        {
                                            try
                                            {
                                                agroup.Members.Add(UserPrincipal.FindByIdentity(context, user.Trim()));
                                            }
                                            catch (Exception ex)
                                            {
                                                ErrorList = ErrorList + ex.Message + $" User: \"{user}\" for {group}" + '\n';
                                            }
                                        }
                                        else if (AddorRemove.ToLower() == "remove")
                                        {
                                            try
                                            {
                                                agroup.Members.Remove(UserPrincipal.FindByIdentity(context, user.Trim()));
                                            }
                                            catch (Exception ex)
                                            {
                                                ErrorList = ErrorList + ex.Message + $" User: \"{user}\" for {group}" + '\n';
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ErrorList = ErrorList + ex.Message + $" User: \"{user}\" for {group}" + '\n';
                                    }

                                    try
                                    {
                                        agroup.Save();
                                    }
                                    catch (Exception ex)
                                    {
                                        ErrorList = ErrorList + ex.Message + $" User: \"{user}\" for {group}" + '\n';
                                    }
                                }
                            }
                        }
                    }
                }
                CloseLoadingWindow(Window);
                if (ErrorList != "")
                {
                    MessageBox.Show(ErrorList, "Error", MessageBoxButton.OK, MessageBoxImage.Hand);
                    System.Windows.Clipboard.SetText(ErrorList);
                }
                MessageBox.Show("All user(s)/group(s) have been processed", "Processing", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                MessageBox.Show("Input Missing", "Error", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        private void RemoveUGButton_Click(object sender, RoutedEventArgs e)
        {
            string usertext = StringFromRTB(ARusernameBox);
            string grouptext = StringFromRTB(ARgroupBox);
            string[] myUserList = usertext.Split("\r\n").SkipLast(1).ToArray();
            string[] myGroupList = grouptext.Split("\r\n").SkipLast(1).ToArray();
            AddOrRemoveBulk(myUserList, myGroupList, "remove");
        }

        private void AddUGButton_Click(object sender, RoutedEventArgs e)
        {
            string usertext = StringFromRTB(ARusernameBox);
            string grouptext = StringFromRTB(ARgroupBox);
            string[] myUserList = usertext.Split("\r\n").SkipLast(1).ToArray();
            string[] myGroupList = grouptext.Split("\r\n").SkipLast(1).ToArray();
            AddOrRemoveBulk(myUserList, myGroupList, "add");
        }

        private void AddGP_Click(object sender, RoutedEventArgs e)
        {
            string usertext = StringFromRTB(ARusernameBox);
            string[] myUserList = usertext.Split("\r\n").SkipLast(1).ToArray();
            string[] myGroupList = { "ISDU_VPN_GP_FullAccess" };
            AddOrRemoveBulk(myUserList, myGroupList, "add");
        }

        private void AddCisco_Click(object sender, RoutedEventArgs e)
        {
            string usertext = StringFromRTB(ARusernameBox);
            string[] myUserList = usertext.Split("\r\n").SkipLast(1).ToArray();
            string[] myGroupList = { "ISDG_VPN_FullAccess" };
            AddOrRemoveBulk(myUserList, myGroupList, "add");
        }

        private void AddMyApps_Click(object sender, RoutedEventArgs e)
        {
            string usertext = StringFromRTB(ARusernameBox);
            string[] myUserList = usertext.Split("\r\n").SkipLast(1).ToArray();
            string[] myGroupList = { "ISDU_CitrixAccessGateway" };
            AddOrRemoveBulk(myUserList, myGroupList, "add");
        }

        private void AddeRecord_Click(object sender, RoutedEventArgs e)
        {
            string usertext = StringFromRTB(ARusernameBox);
            string[] myUserList = usertext.Split("\r\n").SkipLast(1).ToArray();
            string[] myGroupList = { "ISDG_CTX_eRecord2" };
            AddOrRemoveBulk(myUserList, myGroupList, "add");
        }
    }
}