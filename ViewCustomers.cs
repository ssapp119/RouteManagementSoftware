using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Principal;


namespace UpdateCustomers
{
    public partial class ViewCustomers : Form
    {
        public bool NameChange = false;
        public bool SelectedValueChanged = false;
        public string prevName;
        public string prevDescription;
        public string prevStreet;
        public string prevCity;
        public string prevState;
        public string prevZip;
        public string prevNotes;
        public string prevPrice;
        public int items;


        public List<Routes> routelist;// = new List<Routes>();
        public List<Routes> loadedroutelist;

        private static List<Customer> customerlist;
        public static List<Driver> driverlist;
        public int currentindexofsearchlist = 0;
        public List<int> searchlist;

        public ViewCustomers()
        {
            
            InitializeComponent();


            btnUpdateCustomer.Enabled = false;
            btnDelete.Enabled = false;
            btnClear.Enabled = false;
            radAsc.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            radDesc.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            radNone.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            //       GrantAccess((string)Properties.Settings.Default["Path"] + "\\PrintingSoftware");
          //  BringInRouteList();

          //  BringInDriverList();
           // UpdateCustomerListview();
            //  MatchupRouteandCustomerFiles();
            //  MatchupDriverandRouteFiles();
            loadCustomersToolStripMenuItem.Visible = false;
            lstCustomer.View = View.List;
            lstCustomer.View = View.Details;
            lstCustomer.FullRowSelect = true;
            //lstCustomer.Sorting = SortOrder.Ascending;
            txtNotes.MaxLength = 80;
            txtState.MaxLength = 15;
            txtzip.MaxLength = 5;
           

            lstCustomer.Location = new Point(100, 400);
            lstCustomer.Size = new Size(1200, 200);

            lblTotal.Location = new Point(100, 650);
            lblTotalAmt.Location = new Point(200, 650);


            //FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Maximized;

            //Properties.Settings.Default["Activated"] = false;
            //if ((bool)Properties.Settings.Default["FirstRun"] == false)
            //{
            //    using (ProductCodeAuthorization pca = new ProductCodeAuthorization())
            //    {

            //    }

            //}

                    //Properties.Settings.Default["FirstRun"] = true;
            if ((bool)Properties.Settings.Default["FirstRun"] == true)
            {

                using (FolderPath itemcount = new FolderPath())
                {
                    Properties.Settings.Default.Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    //GrantAccess((string)Properties.Settings.Default["Path"] + "\\PrintingSoftware");
                    //if (!Directory.Exists((string)Properties.Settings.Default["Path"] + "\\RouteManagementFiles"))
                    //    Directory.CreateDirectory((string)Properties.Settings.Default["Path"] + "\\RouteManagementFiles");

                    Properties.Settings.Default.FirstRun = false;
                    Properties.Settings.Default.Save();
                    // do what ever with result...
                    itemcount.Text = "Route Management Software - v 1.0.0.50";
                    var result = itemcount.ShowDialog();
                    //if (result == DialogResult.OK)
                    //{
                    //    Properties.Settings.Default["Path"] = itemcount.Path;
                    //    // countresults = itemcount.Count;


                    //}
                    //else
                    //{
                    //    // return;
                    //}

                }

            }

            lstCustomer.Items.Clear();
            BringInDriverList();
            BringInRouteList();
            UpdateCustomerListview();
            //   GetPath();
            lblTotalAmt.Text = items.ToString();
            //txtFolder.Text = (string)Properties.Settings.Default["Path"];


            // Create the list to use as the custom source. 
            var source = new AutoCompleteStringCollection();

           // var citylist = customerlist.Select(p => p.City);

            foreach (var item in customerlist.Select(p => p.City))
            {
                source.Add(item.ToString());
            }


            txtCity.AutoCompleteCustomSource = source;
            // txtCity.AutoCompleteCustomSource
            // Create and initialize the text box.


            // Add the text box to the form.


        }

        private void radioButtons_CheckedChanged(object sender, EventArgs e)
        {
            if (radAsc.Checked)
            {
                lstCustomer.Sorting = SortOrder.Ascending;
                lstCustomer.SelectedItems.Clear();
                btnUpdateCustomer.Enabled = false;
                btnDelete.Enabled = false;
                txtName.Clear();
                txtDescription.Clear();
                txtStreet.Clear();
                txtState.Clear();
                txtzip.Clear();
                txtNotes.Clear();
                txtPrice.Clear();
                txtCity.Clear();
            }
            if (radDesc.Checked)
            {
                lstCustomer.Sorting = SortOrder.Descending;
                lstCustomer.SelectedItems.Clear();
                btnUpdateCustomer.Enabled = false;
                btnDelete.Enabled = false;
                txtName.Clear();
                txtDescription.Clear();
                txtStreet.Clear();
                txtState.Clear();
                txtzip.Clear();
                txtNotes.Clear();
                txtPrice.Clear();
                txtCity.Clear();
            }
            if (radNone.Checked)
            {
                lstCustomer.Sorting = SortOrder.None;
                //lstCustomer.Clear();
                lstCustomer.Items.Clear();
                UpdateCustomerListFromActualList();

            }
        }

        private void GetPath()
        {
            //Properties.Settings.Default["FirstRun"] = true;
            
            if ((bool)Properties.Settings.Default["FirstRun"] == true)
            {

                using (FolderPath itemcount = new FolderPath())
                {


                    // do what ever with result...
                    itemcount.Text = "Route Management Software - v 1.0.0.50";
                    var result = itemcount.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                       
                        Properties.Settings.Default["Path"] = itemcount.Path;
                        // countresults = itemcount.Count;

                        Properties.Settings.Default.Save();

                    }
                    else
                    {
                      //  return;
                    }

                }

            }

            else
            {
                MessageBox.Show("First run is false");
                //Not first time of running application.
            }
        }



        //private void GrantAccess(string fullPath)
        //{
        //    DirectoryInfo dInfo = new DirectoryInfo(fullPath);
        //    DirectorySecurity dSecurity = dInfo.GetAccessControl();
        //    dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
        //    dInfo.SetAccessControl(dSecurity);
        //}

        private void BringInDriverList()
        {
            routelist = new List<Routes>();
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Drivers.txt";



            driverlist = new List<Driver>();
            if (File.Exists(filepath))
            {
                var lines = File.ReadAllLines(filepath);
                foreach (var line in lines)
                {


                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = (line.ToString());
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);

                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;

                    string weirdcharcters = "|";
                    string[] tokens = decryptedstring.Split(new[] { weirdcharcters }, StringSplitOptions.None);

                    driverlist.Add(new Driver { Name = tokens[0].ToString(), Address = tokens[1].ToString() });

                }
            }
        }


        public class Driver
        {
            public string Name { get; set; }
            public string Address { get; set; }
        }

        //private void MatchupDriverandRouteFiles()
        //{
        //    if (routelist.Count > 0)
        //    {
        //        foreach(var item in routelist)
        //        {
                    
        //        }
        //    }
        //}

        //private void MatchupRouteandCustomerFiles()
        //{

        //    if (routelist.Count > 0)
        //    {
        //        foreach(var itroute in routelist)
        //        {

        //        }
        //    }
            
        //}

        private void BringInRouteList()
        {
            //loadedroutelist = new List<Routes>();
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";



            loadedroutelist = new List<Routes>();
            if (File.Exists(filepath))
            {
                var lines = File.ReadAllLines(filepath);
                foreach (var line in lines)
                {


                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = (line.ToString());
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);

                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;


                    string weirdcharcters = "|";
                    string[] tokens = decryptedstring.Split(new[] { weirdcharcters }, StringSplitOptions.None);

                    loadedroutelist.Add(new Routes { driver = tokens[0].ToString(), day = tokens[1].ToString(), customername = tokens[2].ToString(), street = tokens[3].ToString(), city = tokens[4].ToString(), state = tokens[5].ToString(), zip = tokens[6].ToString(), description = tokens[7].ToString(), notes = tokens[8].ToString(), price = tokens[9].ToString() });

                }
            }
        }
        private void UpdateCustomerListview()
        {
            string line;
             items = 0;
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\CustomerLoad.txt";


            // string[] tokens;
            customerlist = new List<Customer>();
            string FileName = (filepath);
            using (FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite ))
            using (StreamReader sr = new StreamReader(fs))
            {
                while ((line = sr.ReadLine()) != null)
                   // foreach (var line in sr.ReadToEnd())
                {
                    items++;
                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = (line.ToString());
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);

                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;




                   // string strData = line.ToString();

                 //   string weirdcharcters = "|";

                    //     string[] tokens = line.Split(new[] { weirdcharcters }, StringSplitOptions.None);

                    string[] tokens = decryptedstring.Split('|');
                    customerlist.Add(new Customer { Name = tokens[0].ToString(), Description = tokens[1].ToString(), StreetNumber = tokens[2].ToString(), City = tokens[3].ToString(), State = tokens[4].ToString(), Zip = tokens[5].ToString(), Notes = tokens[6].ToString(), Price = tokens[7].ToString() });
                }
            }




            
            //var lines = File.ReadAllLines(@"C:\Users\Owner\Desktop\customerload.txt");
            //foreach (var line in lines)
            //{


            //    string weirdcharcters = "|";
            //    string[] tokens = line.Split(new[] { weirdcharcters }, StringSplitOptions.None);


            //    customerlist.Add(new Customer { Name = tokens[0].ToString(), Description = tokens[1].ToString(), StreetNumber = tokens[2].ToString(), City = tokens[3].ToString(), State = tokens[4].ToString(), Zip = tokens[5].ToString() });

            //}


            foreach (var item in customerlist)
            {
                ListViewItem lstItem = new ListViewItem(item.Name);
                lstItem.SubItems.Add(item.Description);
                lstItem.SubItems.Add(item.StreetNumber);
                lstItem.SubItems.Add(item.City);
                lstItem.SubItems.Add(item.State);
                lstItem.SubItems.Add(item.Zip);
                lstItem.SubItems.Add(item.Notes);
                lstItem.SubItems.Add(item.Price);
                lstCustomer.Items.Add(lstItem);

                
                //  lstFullList.Items.Add(item.Name + "|" + item.StreetNumber);
            }
        }


    

        private void lv_DoubleClick(object sender, MouseEventArgs e)
        {


            //foreach (ListViewItem item in lstCustomer.Items)
            //{

            //    item.Selected = false;
            //    item.BackColor = Color.White;
            //    item.ForeColor = Color.Black;
            //    // item.Remove();

            //}




                SelectedValueChanged = true;
            //CustomerDetail customerdetail = new CustomerDetail();
            //customerdetail.Name = lstCustomer.SelectedItems[0].SubItems[0].Text;
            //customerdetail.Street = lstCustomer.SelectedItems[0].SubItems[1].Text;
            //customerdetail.City = lstCustomer.SelectedItems[0].SubItems[2].Text;
            
            //customerdetail.ShowDialog();
            txtName.Text = lstCustomer.SelectedItems[0].SubItems[0].Text;
            txtDescription.Text = lstCustomer.SelectedItems[0].SubItems[1].Text;
            txtStreet.Text = lstCustomer.SelectedItems[0].SubItems[2].Text;
            txtCity.Text  = lstCustomer.SelectedItems[0].SubItems[3].Text;
            txtState.Text = lstCustomer.SelectedItems[0].SubItems[4].Text;
            txtzip.Text = lstCustomer.SelectedItems[0].SubItems[5].Text;
            txtNotes.Text = lstCustomer.SelectedItems[0].SubItems[6].Text;
            txtPrice.Text = lstCustomer.SelectedItems[0].SubItems[7].Text;

             prevName = lstCustomer.SelectedItems[0].SubItems[0].Text;
             prevDescription = lstCustomer.SelectedItems[0].SubItems[1].Text;
             prevStreet = lstCustomer.SelectedItems[0].SubItems[2].Text;
             prevCity = lstCustomer.SelectedItems[0].SubItems[3].Text;
             prevState = lstCustomer.SelectedItems[0].SubItems[4].Text;
             prevZip = lstCustomer.SelectedItems[0].SubItems[5].Text;
             prevNotes = lstCustomer.SelectedItems[0].SubItems[6].Text;
             prevPrice = lstCustomer.SelectedItems[0].SubItems[7].Text;

            btnClear.Enabled = true;
            btnDelete.Enabled = true;
            btnSaveCustomer.Enabled = false;

            CustomerDetail(lstCustomer.SelectedItems[0].SubItems[0].Text, lstCustomer.SelectedItems[0].SubItems[2].Text, lstCustomer.SelectedItems[0].SubItems[3].Text, lstCustomer.SelectedItems[0].SubItems[4].Text, lstCustomer.SelectedItems[0].SubItems[5].Text);



            //   CustomerDetail customerdetail = new CustomerDetail(name, street, city, state, zip);



            //    itemcount.ShowDialog();

            //    string countresults = itemcount.Count;
            //customerdetail.Street = "this is the street";
            // do what ever with result...

            //        var result = customerdetail.ShowDialog();
            //        if (result == DialogResult.OK)
            //      {
            //customerdetail.State = "IA";
            //returneddate = datepicker.ChosenDate;

            // return true;

            //    }
            //    else
            //     {
            //  return false;
            //   }

            SelectedValueChanged = false;


        }

        public void CustomerDetail(string name, string street, string city, string state, string zip)
        {

            chkMonday.Checked = false;
            chkTuesday.Checked = false;
            chkWednesday.Checked = false;
            chkThursday.Checked = false;
            chkFriday.Checked = false;
            chkSaturday.Checked = false;
            chkSunday.Checked = false;



            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";

            

            routelist = new List<Routes>();
            if (File.Exists(filepath))
            {
                var lines = File.ReadAllLines(filepath);
                foreach (var line in lines)
                {


                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = (line.ToString());
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);


                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;

                    string weirdcharcters = "|";
                    string[] tokens = decryptedstring.Split(new[] { weirdcharcters }, StringSplitOptions.None);

                    if (tokens[2].ToString() == name)
                    {
                        routelist.Add(new Routes { driver = tokens[0].ToString(), day = tokens[1].ToString(), customername = tokens[2].ToString(), street = tokens[3].ToString(), city = tokens[4].ToString(), state = tokens[5].ToString(), zip = tokens[6].ToString(), description = tokens[7].ToString(), notes = tokens[8].ToString(), price = tokens[9].ToString() });
                    }
                }


                foreach (var item in routelist)
                {
                    switch (item.day)
                    {
                        case "Monday":
                            chkMonday.Checked = true;
                            break;
                        case "Tuesday":
                            chkTuesday.Checked = true;
                            break;
                        case "Wednesday":
                            chkWednesday.Checked = true;
                            break;
                        case "Thursday":
                            chkThursday.Checked = true;
                            break;
                        case "Friday":
                            chkFriday.Checked = true;
                            break;
                        case "Saturday":
                            chkSaturday.Checked = true;
                            break;
                        case "Sunday":
                            chkSunday.Checked = true;
                            break;



                            // break;
                    }
                }
            }

        }

        public class Routes
        {
            public string driver { get; set; }
            public string day { get; set; }
            public string customername { get; set; }

            
            public string street { get; set; }

            public string city { get; set; }

            public string state { get; set; }

            public string zip { get; set; }

            public string description { get; set; }

            public string notes { get; set;}

            public string price { get; set; }
        }
    









    private void btnSaveCustomer_Click(object sender, EventArgs e)
        {
            if (txtName.Text != "")
            {
                var item = customerlist.Where(p => p.Name == txtName.Text).Select(p => p.Name);

                if (item.Count() < 1)
                {

                    customerlist.Add(new Customer { Name = txtName.Text, Description = txtDescription.Text, StreetNumber = txtStreet.Text, City = txtCity.Text, State = txtState.Text, Zip = txtzip.Text, Notes = txtNotes.Text, Price = txtPrice.Text });






                    UpdateCustomerTextFile();
                    lstCustomer.Items.Clear();
                    UpdateCustomerListFromActualList();

                    lstCustomer.AutoResizeColumn(0,
        ColumnHeaderAutoResizeStyle.HeaderSize);
                    lstCustomer.AutoResizeColumn(1,
                        ColumnHeaderAutoResizeStyle.ColumnContent);
                    lstCustomer.AutoResizeColumn(2,
                       ColumnHeaderAutoResizeStyle.ColumnContent);

                    txtName.Clear();
                    txtDescription.Clear();
                    txtStreet.Clear();
                    txtCity.Clear();
                    txtState.Clear();
                    txtzip.Clear();
                    txtNotes.Clear();
                    txtPrice.Clear();

                    btnUpdateCustomer.Enabled = false;

                }
                else
                {
                    MessageBox.Show(string.Format("Customer name:'{0}' is already being used", txtName.Text), "Duplicate Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Please fill in a value for Name.", "No value error", MessageBoxButtons.OK);
            }
        }

        private void UpdateCustomerTextFile()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\CustomerLoad.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }


            //string filepath = (string)Properties.Settings.Default["Path"];

            using (FileStream f = new FileStream(filepath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (StreamWriter s = new StreamWriter(f))
            {
                foreach (var item in customerlist)
                {
                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobeencrypted = (item.Name + "|" + item.Description + "|" + item.StreetNumber + "|" + item.City + "|" + item.State + "|" + item.Zip + "|" + item.Notes + "|" + item.Price);
                    string encryptstring = cipher.EncryptString(stringtobeencrypted);




                    s.WriteLine(encryptstring);


                }


                s.Flush();
                s.Close();

            }

        }

        private void UpdateCustomerListFromActualList()
        {
            items = 0;
            foreach (var item in customerlist)
            {
                items++;
                ListViewItem lstItem = new ListViewItem(item.Name);
                lstItem.SubItems.Add(item.Description);
                lstItem.SubItems.Add(item.StreetNumber);
                lstItem.SubItems.Add(item.City);
                lstItem.SubItems.Add(item.State);
                lstItem.SubItems.Add(item.Zip);
                lstItem.SubItems.Add(item.Notes);
                lstItem.SubItems.Add(item.Price);
                lstCustomer.Items.Add(lstItem);


                //  lstFullList.Items.Add(item.Name + "|" + item.StreetNumber);
            }
            lblTotalAmt.Text = items.ToString();
        }

        private void loadCustomersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadCustomers ld = new LoadCustomers();
            ld.Text = "Route Management Software - v 1.0.0.50";
            ld.ShowDialog();

            UpdateCustomerListview();
        }

        private void printRoutesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PrintRoutes pr = new PrintRoutes();
            pr.Text = "Route Management Software - v 1.0.0.50";
            pr.ShowDialog();
        }

        private void updateRoutesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateRoutes uproutes = new UpdateRoutes();
            uproutes.Text = "Route Management Software - v 1.0.0.50";
            uproutes.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void addRoutesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddRoute ar = new AddRoute();
            ar.Text = "Route Management Software - v 1.0.0.50";
            ar.ShowDialog();

        }

        private void filesLocationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderPath fp = new FolderPath();
            fp.Text = "Route Management Software - v 1.0.0.50";
            fp.ShowDialog();
        }

        private void addDriversToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddDrivers ad = new AddDrivers();
            ad.Text = "Route Management Software - v 1.0.0.50";
            ad.ShowDialog();
        }

        private void btnUpdateCustomer_Click(object sender, EventArgs e)
        {
            //txtName.Text = lstCustomer.SelectedItems[0].SubItems[0].Text;
            //txtDescription.Text = lstCustomer.SelectedItems[0].SubItems[1].Text;
            //txtStreet.Text = lstCustomer.SelectedItems[0].SubItems[2].Text;
            //txtCity.Text = lstCustomer.SelectedItems[0].SubItems[3].Text;
            //txtState.Text = lstCustomer.SelectedItems[0].SubItems[4].Text;
            //txtzip.Text = lstCustomer.SelectedItems[0].SubItems[5].Text;
            //txtNotes.Text = lstCustomer.SelectedItems[0].SubItems[6].Text;
            //txtPrice.Text = lstCustomer.SelectedItems[0].SubItems[7].Text;
            if (NameChange)
            {
                int namelist = customerlist.Where(p => p.Name == txtName.Text).Count();

                if (namelist > 0)
                {
                    MessageBox.Show(string.Format("The updated name:'{0}' already exists in the customer list.",  txtName.Text), "Duplicate Error", MessageBoxButtons.OK);
                    //NameChange = false;
                    return;
                }
            }




            if (lstCustomer.SelectedItems.Count > 0)
            {

            }
            else
            {


                foreach (ListViewItem item in lstCustomer.Items)
                {
                    if (item.SubItems[0].Text == txtName.Text && item.SubItems[1].Text == txtDescription.Text && item.SubItems[2].Text == txtStreet.Text && item.SubItems[3].Text == txtCity.Text && item.SubItems[4].Text == txtState.Text && item.SubItems[5].Text == txtzip.Text && item.SubItems[6].Text == txtNotes.Text && item.SubItems[7].Text == txtPrice.Text)
                    {
                        item.Selected = true;
                    }


                }
            }



            int update = customerlist.Count(w => w.Name == lstCustomer.SelectedItems[0].SubItems[0].Text && w.Description == lstCustomer.SelectedItems[0].SubItems[1].Text && w.StreetNumber == lstCustomer.SelectedItems[0].SubItems[2].Text && w.City == lstCustomer.SelectedItems[0].SubItems[3].Text);

            if (update > 0)
            {

                foreach (var update1 in customerlist.Where(w => w.Name == lstCustomer.SelectedItems[0].SubItems[0].Text && w.Description == lstCustomer.SelectedItems[0].SubItems[1].Text && w.StreetNumber == lstCustomer.SelectedItems[0].SubItems[2].Text && w.City == lstCustomer.SelectedItems[0].SubItems[3].Text))
                {
                    update1.Name = txtName.Text;
                    update1.Description = txtDescription.Text;
                    update1.StreetNumber = txtStreet.Text;
                    update1.City = txtCity.Text;
                    update1.State = txtState.Text;
                    update1.Zip = txtzip.Text;
                    update1.Notes = txtNotes.Text;
                    update1.Price = txtPrice.Text;
                }
            }

            int update2 = loadedroutelist.Count(w => w.customername == lstCustomer.SelectedItems[0].SubItems[0].Text && w.description == lstCustomer.SelectedItems[0].SubItems[1].Text && w.street == lstCustomer.SelectedItems[0].SubItems[2].Text && w.city == lstCustomer.SelectedItems[0].SubItems[3].Text);


            if (update2 > 0)
            {


                foreach (var update3 in loadedroutelist.Where(w => w.customername == lstCustomer.SelectedItems[0].SubItems[0].Text && w.description == lstCustomer.SelectedItems[0].SubItems[1].Text && w.street == lstCustomer.SelectedItems[0].SubItems[2].Text && w.city == lstCustomer.SelectedItems[0].SubItems[3].Text))
                {
                    update3.customername = txtName.Text;
                    update3.description = txtDescription.Text;
                    update3.street = txtStreet.Text;
                    update3.city = txtCity.Text;
                    update3.state = txtState.Text;
                    update3.zip = txtzip.Text;
                    update3.notes = txtNotes.Text;
                    update3.price = txtPrice.Text;
                }
            }
         //   routelist.Where(w => w.customername == "height").ToList().ForEach(s => s.Value = 30);


            UpdateCustomerTextFile();
            UpdateRouteListFile();
            lstCustomer.Items.Clear();
            UpdateCustomerListFromActualList();

            MessageBox.Show("Update Successful", "Success", MessageBoxButtons.OK);

            txtName.Clear();
            txtDescription.Clear();
            txtStreet.Clear();
            txtCity.Clear();
            txtState.Clear();
            txtzip.Clear();
            txtNotes.Clear();
            txtPrice.Clear();

            btnUpdateCustomer.Enabled = false;
            btnDelete.Enabled = false;

        }

 

        private void UpdateRouteListFile()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }


            //string filepath = (string)Properties.Settings.Default["Path"];

            using (FileStream f = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
            using (StreamWriter s = new StreamWriter(f))
            {
                foreach (var item in loadedroutelist)
                {
                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobeencrypted = (item.driver + "|" + item.day + "|" + item.customername + "|" + item.street + "|" + item.city + "|" + item.state + "|" + item.zip + "|" + item.description + "|" + item.notes + "|" + item.price);
                    string encryptstring = cipher.EncryptString(stringtobeencrypted);




                    s.WriteLine(encryptstring);


                }


                s.Flush();
                s.Close();

            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (lstCustomer.SelectedItems.Count > 0)
            {

                DialogResult result = MessageBox.Show("Are you sure you want to remove this customer?  Choosing to remove this customer will also remove this customer from any routes its associated with.  Please proceed with caution.", "Confirmation", MessageBoxButtons.YesNoCancel);

                if (result == DialogResult.Yes)
                {
                    RemoveItemFromRoutes();
                    lstCustomer.SelectedItems[0].Remove();
                    customerlist.Clear();
                    RepopulateListFromListView();

                    txtName.Clear();
                    txtDescription.Clear();
                    txtStreet.Clear();
                    txtCity.Clear();
                    txtState.Clear();
                    txtzip.Clear();
                    txtNotes.Clear();
                    txtPrice.Clear();

                    btnSaveCustomer.Enabled = true;
                    btnDelete.Enabled = false;
                }

                
            }
            else
            {
                MessageBox.Show("Please selected an item.", "No Item Selected", MessageBoxButtons.OK);
            }
            




           
        }

        private void RemoveItemFromRoutes()
        {

            loadedroutelist.RemoveAll(item => item.customername == lstCustomer.SelectedItems[0].SubItems[0].Text && item.description == lstCustomer.SelectedItems[0].SubItems[1].Text
                    && item.street == lstCustomer.SelectedItems[0].SubItems[2].Text && item.city == lstCustomer.SelectedItems[0].SubItems[3].Text
                    && item.state == lstCustomer.SelectedItems[0].SubItems[4].Text && item.zip == lstCustomer.SelectedItems[0].SubItems[5].Text
                    && item.notes == lstCustomer.SelectedItems[0].SubItems[6].Text && item.price == lstCustomer.SelectedItems[0].SubItems[7].Text);

            //foreach (var item in loadedroutelist)
            //{
            //    if (item.customername == lstCustomer.SelectedItems[0].SubItems[0].Text && item.description == lstCustomer.SelectedItems[0].SubItems[1].Text
            //        && item.street == lstCustomer.SelectedItems[0].SubItems[2].Text && item.city == lstCustomer.SelectedItems[0].SubItems[3].Text
            //        && item.state == lstCustomer.SelectedItems[0].SubItems[4].Text && item.zip == lstCustomer.SelectedItems[0].SubItems[5].Text
            //        && item.notes == lstCustomer.SelectedItems[0].SubItems[6].Text && item.price == lstCustomer.SelectedItems[0].SubItems[7].Text)
            //    {
            //        loadedroutelist.Remove(item);
            //    }

            //}

            RepopulateRouteTextFileFromList();
        }



        private void RepopulateRouteTextFileFromList()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }


            //string filepath = (string)Properties.Settings.Default["Path"];

            using (FileStream f = new FileStream(filepath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (StreamWriter s = new StreamWriter(f))
            {
                foreach (var item in loadedroutelist)
                {
                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobeencrypted = (item.driver + "|" + item.day + "|" + item.customername + "|" + item.street + "|" + item.city + "|" + item.state + "|" + item.zip + "|" + item.description + "|" + item.notes + "|" + item.price);
                    string encryptstring = cipher.EncryptString(stringtobeencrypted);




                    s.WriteLine(encryptstring);


                }


                s.Flush();
                s.Close();

            }
        }

        private void RepopulateListFromListView()
        {
            items = 0;
            foreach (ListViewItem itemRow in lstCustomer.Items)
            {
                items++;
                    customerlist.Add(new Customer { Name = itemRow.SubItems[0].Text, Description = itemRow.SubItems[1].Text, StreetNumber = itemRow.SubItems[2].Text, City = itemRow.SubItems[3].Text, State = itemRow.SubItems[4].Text, Zip = itemRow.SubItems[5].Text, Notes = itemRow.SubItems[6].Text, Price = itemRow.SubItems[7].Text });

                //    ListView1.SelectedItems[0].Subitems[1].Text
             }
            lblTotalAmt.Text = items.ToString();
            UpdateCustomerTextFile();

        }

        private void txtSearch_TextChanged_1(object sender, EventArgs e)
        {
            currentindexofsearchlist = 0;
            int values = 0;

            foreach (ListViewItem item in lstCustomer.Items)
            {
                if (txtSearch.Text != "")
                {
                    if (item.Text.ToLower().StartsWith(txtSearch.Text.ToLower()))
                    {
                        item.Selected = true;

                        item.BackColor = Color.CornflowerBlue;
                        item.ForeColor = Color.White;
                        values++;
                    }

                    else
                    {
                        item.Selected = false;
                        item.BackColor = Color.White;
                        item.ForeColor = Color.Black;
                        // item.Remove();
                    }
                }
                else
                {
                    item.Selected = false;
                    item.BackColor = Color.White;
                    item.ForeColor = Color.Black;
                }
                lblSearchItems.Text = values.ToString();

            }
                //if (ListViewCostumControl.lvnf.SelectedItems.Count == 1)
                //{
                //    ListViewCostumControl.lvnf.Focus();
                //}



            searchlist = new List<int>();


            foreach (ListViewItem item in lstCustomer.Items)
            {
                if (item.BackColor == Color.CornflowerBlue)
                    searchlist.Add(item.Index);
            }
            if (searchlist.Count > 0)
                lstCustomer.TopItem = lstCustomer.Items[searchlist.FirstOrDefault()];
            //  lstFullList.TopItem = searchlist.item;
            //  lstFullList.TopItem = lstFullList.item

            //  list.TopItem = list.Item[x]
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (lstCustomer.SelectedItems.Count > 0)
            {
                if (searchlist.Count > 0)
                {
                    if (currentindexofsearchlist > searchlist.Count - 1)
                        currentindexofsearchlist = 0;


                    lstCustomer.TopItem = lstCustomer.Items[searchlist[currentindexofsearchlist]];
                }


                currentindexofsearchlist++;
            }
        }

        private void companyNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CompanyName cn = new CompanyName();
            cn.Text = "Route Management Software - v 1.0.0.50";
            cn.ShowDialog();
        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                   if (prevName != txtName.Text)
                    {
                        NameChange = true;
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                        NameChange = false;
                    }

                }
                
            }


        }

        //private void lstCustomer_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    SelectedValueChanged = true;
        //}

        private void txtDescription_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevDescription != txtDescription.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void txtStreet_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevStreet != txtStreet.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void txtCity_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevCity != txtCity.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void txtState_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevState != txtState.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void txtzip_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevZip != txtzip.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void txtPrice_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevPrice != txtPrice.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void txtNotes_TextChanged(object sender, EventArgs e)
        {
            if (!SelectedValueChanged)
            {
                if (lstCustomer.SelectedItems.Count > 0)
                {
                    if (prevNotes != txtNotes.Text)
                    {
                        btnUpdateCustomer.Enabled = true;
                    }
                    else
                    {
                        btnUpdateCustomer.Enabled = false;
                    }
                }

            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {


                txtName.Clear();
            txtDescription.Clear();
            txtStreet.Clear();
            txtState.Clear();
            txtCity.Clear();
            txtPrice.Clear();
            txtzip.Clear();
            txtNotes.Clear();

            lstCustomer.SelectedItems.Clear();

            btnDelete.Enabled = false;
            btnUpdateCustomer.Enabled = false;


            btnSaveCustomer.Enabled = true;
        }
    }
}
    

