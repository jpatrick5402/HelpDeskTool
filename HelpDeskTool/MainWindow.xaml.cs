/// Program HDTool written by Joseph Patrick for use in the URMC ISD Help Desk environment
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
using Syncfusion.XlsIO;
using System.IO;

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
            MessageBox.Show("No PC Name/IP Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            OutputBox.AppendText("No PC Name/IP Detected");
            return false;
        }
        public Boolean IsTextInUserBox()
        {
            if (UserTextbox.Text != "")
                return true;
            MessageBox.Show("No AD Name Detected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            OutputBox.AppendText("No AD Name Detected");
            return false;
        }
        //Restart Button
        private void RestartButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInNameBox())
            {
                MessageBoxResult result = MessageBox.Show("Are you sure you want to restart this computer?",
                                          "Confirmation",
                                          MessageBoxButton.YesNo,
                                          MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
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
                else
                {
                    OutputBox.AppendText("PC not restarted\n");
                }
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Ping Button
        private void PingButton_Click(object sender, RoutedEventArgs e)
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
        private void SysinfoButton_Click(object sender, RoutedEventArgs e)
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
        private void ADinfoButton_Click(object sender, RoutedEventArgs e)
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
        private void MemberOfButton_Click(object sender, RoutedEventArgs e)
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
                command.StartInfo.Arguments = "Get-ADPrincipalGroupMembership " + Username + " | select name | Sort-Object -Property name";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Get Serial Button
        private void SerialButton_Click(object sender, RoutedEventArgs e)
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
        private void MembersButton_Click(object sender, RoutedEventArgs e)
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
                command.StartInfo.Arguments = "Get-ADGroupMember -Identity \'" + Username + "\' | select SamAccountName | Sort-Object -Property SamAccountName";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd().Replace(" ", ""));
                
                
                OutputBox.AppendText("\n\n");

                // Gather names
                System.Diagnostics.Process command2 = new System.Diagnostics.Process();
                command2.StartInfo.CreateNoWindow = true;
                command2.StartInfo.FileName = "powershell";
                command2.StartInfo.Arguments = "Get-ADGroupMember -Identity \'" + Username + "\' | select name | Sort-Object -Property name";
                command2.StartInfo.RedirectStandardOutput = true;
                command2.Start();
                OutputBox.AppendText(command2.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        //Add group button
        private void AddgroupButton_Click(object sender, RoutedEventArgs e)
        {
            AddRemoveWindow win = new AddRemoveWindow();
            win.Show();
        }
        // Pass Last Set Button
        // Runs powershell commands to get AD info which includes the user's Password set date and when it last failed
        // There is a current issue determining the accuracy of the Last Bad Password fucntion because there are more that one DC's to use
        private void PassButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsTextInUserBox())
            {
                OutputBox.AppendText("\nLastBadPasswordAttempt NOT ALWAYS ACCURATE\n");

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

            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }
        // Clear Button
        // Completely deletes all text from the main text box
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            OutputBox.Document.Blocks.Clear();
        }

        private void ADGroupInfoButton_Click(object sender, RoutedEventArgs e)
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
                command.StartInfo.Arguments = "Get-ADGroup \'" + Username + "\' -Properties *";
                command.StartInfo.RedirectStandardOutput = true;
                command.Start();
                OutputBox.AppendText(command.StandardOutput.ReadToEnd());
            }
            OutputBox.AppendText("\n---------------------------------------------------------------------------------------------------------------------------------------\n");
            OutputBox.ScrollToEnd();
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            TextRange textRange = new(
                OutputBox.Document.ContentStart,
                OutputBox.Document.ContentEnd
            );
            ExcelEngine excelEngine = new ExcelEngine();
            var text = textRange.Text;
            List<string> separatedList = new(text.Split("\r\n"));
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            //Reads input Excel stream as a workbook
            IWorkbook workbook = application.Workbooks.Open(File.OpenRead(System.IO.Path.GetFullPath(@"C:\\Users\\USER\Desktop\Expenses.xlsx")));
            IWorksheet sheet = workbook.Worksheets[0];

            //Preparing first array with different data types
            object[] expenseArray = new object[14]
            {"Paul Pogba", 469.00d, 263.00d, 131.00d, 139.00d, 474.00d, 253.00d, 467.00d, 142.00d, 417.00d, 324.00d, 328.00d, 497.00d, "=SUM(B11:M11)"};

            //Inserting a new row by formatting as a previous row.
            sheet.InsertRow(11, 1, ExcelInsertOptions.FormatAsBefore);

            //Import Peter's expenses and fill it horizontally
            sheet.ImportArray(expenseArray, 11, 1, false);

            //Preparing second array with double data type
            double[] expensesOnDec = new double[6]
            {179.00d, 298.00d, 484.00d, 145.00d, 20.00d, 497.00d};

            //Modify the December month's expenses and import it vertically
            sheet.ImportArray(expensesOnDec, 6, 13, true);

            //Save the file in the given path
            Stream excelStream = File.Create(System.IO.Path.GetFullPath(@"Output.xlsx"));
            workbook.SaveAs(excelStream);
            excelStream.Dispose();
        }
    }
}
