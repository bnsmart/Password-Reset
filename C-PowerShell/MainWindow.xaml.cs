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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management.Automation;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.Data;
using MahApps.Metro.Controls;
using ClosedXML.Excel;
using System.ComponentModel;
using System.Threading;
using System.IO;

namespace C_PowerShell
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public string NonAd;
        public string schooltype;
        public string schoolname;
        public string usertype1;
        public string usertype2;
        public string YOLO;
        public string FinalQuery;
        DataTable Results = new DataTable();

        private readonly BackgroundWorker worker = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();
            PWreset.IsEnabled = false;
            richTextBox.Document.Blocks.Clear();
            Results.Columns.Add("First Name");
            Results.Columns.Add("Last Name");
            Results.Columns.Add("User Name");
            Results.Columns.Add("Password");
            Results.Columns.Add("Result");
            userdataGrid.ItemsSource = Results.DefaultView;

            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.ProgressChanged += worker_ProgressChanged;

            Schooltype_comboBox.Items.Add("Primary Schools");
            Schooltype_comboBox.Items.Add("PRU");
            Schooltype_comboBox.Items.Add("Misc. Centres");
            Schooltype_comboBox.Items.Add("Special Schools");
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (textBox1.Text != "")
            {
                PowerShell PowerShellInstance = PowerShell.Create();

                richTextBox.Document.Blocks.Clear();
                Cursor = Cursors.Wait;
                string name = textBox1.Text;

                System.IO.Directory.CreateDirectory(@"C:\Users\" + Environment.UserName + @"\Documents\Powershell_Reports");

                PowerShellInstance.AddScript("Get - Job | Remove - Job");
                PowerShellInstance.AddScript("$computers = Get-ADComputer -ResultSetSize 100 -Filter \"name -like \"*" + name + "*\"\" | select -expand name");

                if (radioButton_serial.IsChecked == true)
                {
                    PowerShellInstance.AddScript(@"$output = foreach ($computer in $computers) { If (Test-Connection $computer -count 1 -quiet) { Get-WmiObject win32_bios -computername $computer | Select-Object __Server, SerialNumber, Manufacturer, Name, Version -ErrorAction Stop } }");
                    PowerShellInstance.AddScript(@"$output | Export-Csv -Path $env:userprofile\Documents\Powershell_Reports\Serials.csv -Encoding ascii -NoTypeInformation");
                }
                else if (radioButton_services.IsChecked == true)
                {
                    PowerShellInstance.AddScript("$computers = Get-ADComputer -ResultSetSize 1 -Filter \"name -like \"*" + name + "*\"\" | select -expand name");
                    PowerShellInstance.AddScript(@"$output = foreach ($computer in $computers) { If (Test-Connection $computer -count 1 -quiet) { Get-Service -computername $computer } }");
                    PowerShellInstance.AddScript(@"$output | Export-Csv -Path $env:userprofile\Documents\Powershell_Reports\Services.csv -Encoding ascii -NoTypeInformation");
                }
                else if (radioButton_lastLogon.IsChecked == true)
                {
                    PowerShellInstance.AddScript(@"$output = foreach ($computer in $computers) { Get-ADComputer $computer -Properties LastLogonTimeStamp | select-object name, @{Name=""Last Logon""; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} }");
                    PowerShellInstance.AddScript(@"$output | Export-Csv -Path $env:userprofile\Documents\Powershell_Reports\Last_Logon.csv -Encoding ascii -NoTypeInformation");
                }
                else if (Sccm_Client_Version.IsChecked == true)
                {
                    PowerShellInstance.AddScript(@"$cmd = { param($computer) If (Test-Connection $computer -count 1 -quiet) { Get-WMIObject -namespace root\ccm -class sms_client -computername $computer } }");
                    PowerShellInstance.AddScript(@"$jobs = foreach ($computer in $computers) { Start-Job -ScriptBlock $cmd -ArgumentList $computer | Out-Null }");
                    PowerShellInstance.AddScript(@"Get-Job | Wait-Job");
                    PowerShellInstance.AddScript(@"$output = Get-Job | Receive-Job | Select-Object __Server, ClientVersion");
                    PowerShellInstance.AddScript(@"$output | Export-Csv -Path $env:userprofile\Documents\Powershell_Reports\SCCM.csv -Encoding ascii -NoTypeInformation");
                }
                else
                {
                    PowerShellInstance.AddScript("Write-Output \"Nothing to selected.\"");
                };
                PowerShellInstance.AddScript(@"$output | out-gridview");
                PowerShellInstance.AddScript(@"$output | ft -AutoSize | Out-String");

                var result = PowerShellInstance.Invoke();
                foreach (var item in result)
                {
                    richTextBox.AppendText(item.ToString());
                }
                Cursor = Cursors.Arrow;
            }
            else
            {
                MessageBox.Show("Text Field Empty");
            }
        }

        public DirectorySearcher QureyAD(string path, string filter, string sort)
        {
            DirectoryEntry solgridAD = new DirectoryEntry(path);

            DirectorySearcher searcher = new DirectorySearcher(solgridAD);
         
            searcher.SearchScope = SearchScope.OneLevel;
            
            searcher.Filter = filter;
            searcher.Sort.PropertyName = sort;
            return searcher;
        }

        private void NonAdcheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (Schooltype_comboBox != null) { Schooltype_comboBox.Items.Clear(); }
            NonAd = ",OU=Non-Ad Schools";
            Schooltype_comboBox.Items.Add("Exchange Users");
            Schooltype_comboBox.Items.Add("Extranet Users");
        }

        private void NonAdcheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (Schooltype_comboBox != null) { Schooltype_comboBox.Items.Clear(); }
            NonAd = "";

            Schooltype_comboBox.Items.Add("Primary Schools");
            Schooltype_comboBox.Items.Add("PRU");
            Schooltype_comboBox.Items.Add("Misc. Centres");
            Schooltype_comboBox.Items.Add("Special Schools");
        }

        private void Schooltype_comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SchoolOUcomboBox != null) { SchoolOUcomboBox.Items.Clear(); }

            getuserButton.IsEnabled = false;
            YOLOUcomboBox.IsEnabled = false;

            if (Schooltype_comboBox.Items.Count > 0)
            {

                schooltype = Schooltype_comboBox.SelectedItem.ToString();

                DirectorySearcher searcher = QureyAD("LDAP://OU=" + schooltype + NonAd + ",DC=solgrid,DC=local", "(&(objectCategory=organizationalUnit)(!ou=Roaming)(!ou=SENTEST)(!ou=Exam user))", "OU");
                FinalQuery = schooltype + NonAd;
                try
                {
                    foreach (SearchResult res in searcher.FindAll())
                    {
                        SchoolOUcomboBox.Items.Add(res.Properties["ou"][0].ToString());
                    }
                }
                catch
                {
                    MessageBox.Show("Can't connect to the Active Directory");
                }
            }

            //IsPrimary();
        }

        private void SchoolOUcomboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (UserTypecomboBox != null) { UserTypecomboBox.Items.Clear(); }

            getuserButton.IsEnabled = false;
            YOLOUcomboBox.IsEnabled = false;

            if (SchoolOUcomboBox.Items.Count > 0)
            {

                schoolname = SchoolOUcomboBox.SelectedItem.ToString();

                DirectorySearcher searcher = QureyAD("LDAP://OU=" + schoolname + ",OU=" + schooltype + NonAd + ",DC=solgrid,DC=local", "(&(objectCategory=organizationalUnit)(!ou=Roaming)(!ou=SENTEST)(!ou=Exam user))", "OU");
                FinalQuery = schoolname + ",OU = " + schooltype + NonAd;
                try
                {
                    foreach (SearchResult res in searcher.FindAll())
                    {

                        if (res.Properties["ou"][0].ToString() == "Admin Accounts")
                        {
                            UserTypecomboBox.Items.Add("Admin Accounts");
                            UserTypecomboBox.Items.Add("Governor Accounts");
                        }
                        else if (res.Properties["ou"][0].ToString() == "Standard Accounts")
                        {
                            UserTypecomboBox.Items.Add("Pupils");
                            UserTypecomboBox.Items.Add("Staff");
                        }

                    }

                }
                catch
                {
                    MessageBox.Show("Can't connect to the Active Directory");
                }

                if (schoolname == "Home Teaching")
                {
                    UserTypecomboBox.IsEnabled = false;
                    getuserButton.IsEnabled = true;
                }
                else
                {
                    UserTypecomboBox.IsEnabled = true;
                    getuserButton.IsEnabled = false;
                }
                //IsPrimary();
            }
        }

        private void UserTypecomboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (YOLOUcomboBox != null) { YOLOUcomboBox.Items.Clear(); }

            YOLOUcomboBox.IsEnabled = false;

            if (UserTypecomboBox.Items.Count > 0)
            {

                usertype2 = UserTypecomboBox.SelectedItem.ToString();

                if (usertype2 == "Admin Accounts")
                {
                    usertype1 = "";
                    FinalQuery = usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd;
                    getuserButton.IsEnabled = true;
                }
                else if (usertype2 == "Governor Accounts")
                {
                    usertype1 = ",OU=Admin Accounts";
                    FinalQuery = usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd;
                    getuserButton.IsEnabled = true;
                }
                else if (usertype2 == "Staff")
                {
                    usertype1 = ",OU=Standard Accounts";
                    FinalQuery = usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd;
                    getuserButton.IsEnabled = true;
                }
                else
                {
                    usertype1 = ",OU=Standard Accounts";
                    FinalQuery = usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd;
                    getuserButton.IsEnabled = true;

                    if (schooltype == "Primary Schools" || NonAd == ",OU=Non-Ad Schools") {
                        YOLOUcomboBox.IsEnabled = true;
                    }
                    DirectorySearcher searcher = QureyAD("LDAP://OU=" + usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd + ",DC=solgrid,DC=local", "(&(objectCategory=organizationalUnit)(!ou=Roaming)(!ou=SENTEST)(!ou=Exam user))", "OU");

                    FinalQuery = usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd;

                    try
                    {
                        foreach (SearchResult res in searcher.FindAll())
                        {
                            YOLOUcomboBox.Items.Add(res.Properties["ou"][0].ToString());
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Can't connect to the Active Directory");
                    }
                }
            }
        }

        private void YOLOUcomboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (YOLOUcomboBox.Items.Count > 0)
            {
                YOLO = YOLOUcomboBox.SelectedItem.ToString();
                FinalQuery = YOLO + ",OU=" + usertype2 + usertype1 + ",OU=" + schoolname + ",OU=" + schooltype + NonAd;
            }
        }

        //public void IsPrimary()
        //{
        //    try
        //    {
        //        if (((ComboBoxItem)Schooltype_comboBox.SelectedItem).Content.ToString() == "Primary Schools")
        //        {
        //            enableButton_primary(((ComboBoxItem)Schooltype_comboBox.SelectedItem).Content.ToString());
        //            //pupilsInou.IsEnabled = true;
        //        }
        //        else
        //        {
        //            enableButton();
        //            //pupilsInou.IsEnabled = false;
        //        }
        //    }
        //    catch
        //    {

        //    }
        //}

        //public void enableButton_primary(string schooltype)
        //{
        //    if (Schooltype_comboBox.SelectedItem != null &&
        //        SchoolOUcomboBox.SelectedItem != null &&
        //        //pupilsInou.IsChecked == true &&
        //        YOLOUcomboBox.SelectedItem == null
        //        )
        //    {        
        //        getuserButton.IsEnabled = false;

        //        YOLOUcomboBox.IsEnabled = true;
        //        YOLOUcomboBox.Items.Clear();
        //        DirectorySearcher searcher = QureyAD("LDAP://OU=Pupils,OU=Standard Accounts,OU=" + SchoolOUcomboBox.SelectedItem.ToString() + ",OU=" + schooltype + ", DC=solgrid,DC=local", "(&(objectCategory=organizationalUnit)(!ou=Roaming)(!ou=SENTEST)(!ou=Exam user))", "OU",2);

        //        try
        //        {
        //            foreach (SearchResult res in searcher.FindAll())
        //            {
        //                YOLOUcomboBox.Items.Add(res.Properties["ou"][0].ToString());
        //            }
        //        }
        //        catch
        //        {
        //            MessageBox.Show("Can't connect to the Active Directory in enableButton_primary");
        //        }

        //        YOLOUcomboBox.IsEnabled = true;
        //    }
        //    else if (Schooltype_comboBox.SelectedItem != null &&
        //        SchoolOUcomboBox.SelectedItem != null
        //        //pupilsInou.IsChecked == false
        //        )
        //    {
        //        getuserButton.IsEnabled = true;
        //    }
        //    else if (SchoolOUcomboBox.SelectedItem != null &&
        //        YOLOUcomboBox.SelectedItem != null)
        //    {
        //        getuserButton.IsEnabled = true;
        //        //MessageBox.Show("YOLO change");
        //    }
        //    else
        //    {
        //        getuserButton.IsEnabled = false;
        //    }
        //}

        //public void enableButton()
        //{
        //    if (Schooltype_comboBox.SelectedItem != null &&
        //        SchoolOUcomboBox.SelectedItem != null
        //        //Govenors.IsChecked == true ||
        //        //allPupils.IsChecked == true ||
        //        //Staff.IsChecked == true
        //        )
        //    {
        //        getuserButton.IsEnabled = true;
        //    }
        //    else if (SchoolOUcomboBox.SelectedItem != null &&
        //        YOLOUcomboBox.SelectedItem != null)
        //    {
        //        getuserButton.IsEnabled = true;
        //        //MessageBox.Show("YOLO change");
        //    }
        //    else
        //    {
        //        getuserButton.IsEnabled = false;
        //    }
        //}

        private void newuserbutton_Click(object sender, RoutedEventArgs e)
        {
            controlContainer.IsEnabled = false;
            senddata();
        }

        public void senddata()
        {

            DirectorySearcher searcher = QureyAD("LDAP://OU=" + FinalQuery + ",DC=solgrid,DC=local", "(objectClass=user)", "sAMAccountName");



            Results.Clear();

            try
            {
                foreach (SearchResult sr in searcher.FindAll())
                {
                    DataRow dr = Results.NewRow();
                    DirectoryEntry de = sr.GetDirectoryEntry();
                    dr["First Name"] = de.Properties["givenName"].Value;
                    dr["Last Name"] = de.Properties["SN"].Value;
                    dr["User Name"] = de.Properties["sAMAccountName"].Value;
                    Results.Rows.Add(dr);
                    de.Close();
                }
            }
            catch
            {
                MessageBox.Show("Can't connect to the Active Directory");
            }

            userdataGrid.ItemsSource = Results.DefaultView;
            controlContainer.IsEnabled = true;
            PWreset.IsEnabled = true;
            PWreset.Focus();
        }

        private void ChangePassBtn_Click_1(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to change password to OU=" + FinalQuery , "Change Passwords", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
            {
                pbStatus.Value = 0;
                worker.RunWorkerAsync();
                controlContainer.IsEnabled = false;
            }

        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {        

                int i = 0;
                foreach (DataRowView com in userdataGrid.Items)
                {

                    DataRow row = com.Row;
                    string userName = row.ItemArray[2].ToString();
                    string passWord = row.ItemArray[3].ToString();

                    if (passWord == "")
                    {
                        Results.Rows[i][4] = "Skipped: No password supplied";
                    }
                    else if (passWord.Length < 6)
                    {
                        Results.Rows[i][4] = "Failed: password Must be 6+ characters";
                    }
                    else
                    {
                        using (var context = new PrincipalContext(ContextType.Domain))
                        {
                            using (var user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, userName))
                            {
                                if (user != null)
                                {
                                    user.SetPassword(passWord);
                                    Results.Rows[i][4] = "Success!!!!!";
                                }
                                else
                                {
                                    Results.Rows[i][4] = "Failed: User does not exist";

                                }
                            }
                        }
                    }

                    i = i + 1;
                    (sender as BackgroundWorker).ReportProgress(i);

                }        

        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            userdataGrid.ItemsSource = Results.DefaultView;

            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(Results, "Worksheet1");
            MessageBox.Show(SchoolOUcomboBox.SelectedItem.ToString() + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx has been saved to " + Directory.GetCurrentDirectory());
            controlContainer.IsEnabled = true;
            string fileName = SchoolOUcomboBox.SelectedItem.ToString() + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            wb.SaveAs(fileName);
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbStatus.Value = (((double)e.ProgressPercentage / (double)userdataGrid.Items.Count) * 100);
        }

        private void Tabs1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        //private void PopulateBtn_Click_1(object sender, RoutedEventArgs e)
        //{
        //    string[] PassWords = {"dad123",
        //                                "pat123",
        //                                "man123",
        //                                "sad123",
        //                                "tap123",
        //                                "sit123",
        //                                "nip123",
        //                                "dim123",
        //                                "sap123",
        //                                "sat123",
        //                                "pan123",
        //                                "mat123",
        //                                "dip123",
        //                                "pit123",
        //                                "pin123",
        //                                "map123",
        //                                "din123",
        //                                "tip123",
        //                                "tin123",
        //                                "did123",
        //                                "pip123",
        //                                "tan123",
        //                                "sip123",
        //                                "nap123",
        //                                "and123",
        //                                "tag123",
        //                                "got123",
        //                                "can123",
        //                                "kid123",
        //                                "gag123",
        //                                "one123",
        //                                "cot123",
        //                                "kit123",
        //                                "gig123",
        //                                "not123",
        //                                "cop123",
        //                                "gap123",
        //                                "pot123",
        //                                "cap123",
        //                                "nag123",
        //                                "top123",
        //                                "cat123",
        //                                "sag123",
        //                                "dog123",
        //                                "cod123",
        //                                "gas123",
        //                                "pop123",
        //                                "pig123",
        //                                "dig123",
        //                                "get123",
        //                                "rim123",
        //                                "pet123",
        //                                "mum123",
        //                                "rip123",
        //                                "ten123",
        //                                "run123",
        //                                "ram123",
        //                                "net123",
        //                                "mug123",
        //                                "rat123",
        //                                "pen123",
        //                                "cup123",
        //                                "rag123",
        //                                "peg123",
        //                                "sun123",
        //                                "rug123",
        //                                "met123",
        //                                "rot123",
        //                                "men123",
        //                                "mud123",
        //                                "sun123",
        //                                "set123",
        //                                "car123",
        //                                "had123",
        //                                "but123",
        //                                "lap123",
        //                                "him123",
        //                                "big123",
        //                                "let123",
        //                                "jot123",
        //                                "his123",
        //                                "off123",
        //                                "leg123",
        //                                "hat123",
        //                                "hot123",
        //                                "bet123",
        //                                "fit123",
        //                                "lot123",
        //                                "map123",
        //                                "hut123",
        //                                "bad123",
        //                                "fin123",
        //                                "lit123",
        //                                "mat123",
        //                                "hop123",
        //                                "bag123",
        //                                "fun123",
        //                                "bee123",
        //                                "bon123",
        //                                "hum123",
        //                                "bed123",
        //                                "fig123",
        //                                "fin123",
        //                                "fun123",
        //                                "hit123",
        //                                "bud123",
        //                                "fog123",
        //                                "dot123",
        //                                "hip123",
        //                                "hat123",
        //                                "beg123",
        //                                "put123",
        //                                "tin123",
        //                                "pan123",
        //                                "has123",
        //                                "bug123",
        //                                "hut123",
        //                                "sew123",
        //                                "kin123",
        //                                "had123",
        //                                "bun123",
        //                                "cup123",
        //                                "hug123",
        //                                "bus123",
        //                                "fan123",
        //                                "fat123",
        //                                "den123",
        //                                "bat123",
        //                                "pot123",
        //                                "bit123"};

        //    int i = 0;
        //    int arraycounter = 0;
        //    foreach (DataRowView com in userdataGrid.Items)
        //    {
        //        Results.Rows[i][1] = PassWords[arraycounter];
        //        arraycounter = arraycounter + 1;
        //        i = i + 1;
        //        if (arraycounter == PassWords.Length) { arraycounter = 0; }
        //    }
        //}

        private void PopulateBtn_Click_1(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true) { }

            string filepath = openFileDialog.FileName;
            try
            {
                string[][] PassWords = File.ReadLines(filepath).Where(line => line != "").Select(x => x.Split('|')).ToArray();
                int i = 0;
                int arraycounter = 0;
                foreach (DataRowView com in userdataGrid.Items)
                {
                    Results.Rows[i][3] = PassWords[arraycounter][0];
                    arraycounter = arraycounter + 1;
                    i = i + 1;
                    if (arraycounter == PassWords.Length) { arraycounter = 0; }
                }
            }
            catch { }


        }

        private void ClearPW_Click_1(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (DataRowView com in userdataGrid.Items)
            {
                Results.Rows[i][3] = "";
                i = i + 1;
            }
        }

    }
}