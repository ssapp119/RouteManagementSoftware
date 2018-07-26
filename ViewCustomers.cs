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
        #region Variables
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
        #endregion

        #region Constructor
        public ViewCustomers()
        {

            InitializeComponent();


            btnUpdateCustomer.Enabled = false;
            btnDelete.Enabled = false;
            btnClear.Enabled = false;
            radAsc.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            radDesc.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            radNone.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            loadCustomersToolStripMenuItem.Visible = false;
            lstCustomer.View = View.List;
            lstCustomer.View = View.Details;
            lstCustomer.FullRowSelect = true;
            txtNotes.MaxLength = 80;
            txtState.MaxLength = 15;
            txtzip.MaxLength = 5;


            lstCustomer.Location = new Point(100, 400);
            lstCustomer.Size = new Size(1200, 200);

            lblTotal.Location = new Point(100, 650);
            lblTotalAmt.Location = new Point(200, 650);


            //FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Maximized;


            //Properties.Settings.Default["FirstRun"] = true;
            if ((bool)Properties.Settings.Default["FirstRun"] == true)
            {

                using (FolderPath itemcount = new FolderPath())
                {
                    Properties.Settings.Default.Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    Properties.Settings.Default.FirstRun = false;
                    Properties.Settings.Default.Save();
                    // do what ever with result...
                    itemcount.Text = "Route Management Software - v 1.0.0.50";
                    var result = itemcount.ShowDialog();


                }

            }

            lstCustomer.Items.Clear();
            BringInDriverList();
            BringInRouteList();
            UpdateCustomerListview();
            //   GetPath();
            lblTotalAmt.Text = items.ToString();



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
        #endregion

        #region Radio Button Changed
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
        #endregion

        #region Get Path
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
        #endregion


        #region DriverList

        private void BringInDriverList()
        {
            routelist = new List<Routes>();
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Drivers.txt";

            driverlist = GetDriverList(filepath);


        }

        #endregion

        #region GetDriverList
        public List<Driver> GetDriverList(string filepath)
        {
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

            return driverlist;


        }
        #endregion

        #region DriverClass
        public class Driver
        {
            public string Name { get; set; }
            public string Address { get; set; }
        }
        #endregion

        #region BringInRouteList
        private void BringInRouteList()
        {
            //loadedroutelist = new List<Routes>();
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";

            loadedroutelist = GetRouteList(filepath);// new List<Routes>();

        }

        private List<Routes> GetRouteList(string filepath)
        {
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
            return loadedroutelist;
        }

        #endregion


        #region UpdateCustomerListView
        private void UpdateCustomerListview()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\CustomerLoad.txt";
            customerlist = GetCustomerList(filepath);// new List<Customer>();

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

            }
        }

        #endregion

        #region GetCustomerList
        public List<Customer> GetCustomerList(string filepath)
        {
            customerlist = new List<Customer>();
            string line;
            items = 0;
            using (FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
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


                    string[] tokens = decryptedstring.Split('|');
                    customerlist.Add(new Customer { Name = tokens[0].ToString(), Description = tokens[1].ToString(), StreetNumber = tokens[2].ToString(), City = tokens[3].ToString(), State = tokens[4].ToString(), Zip = tokens[5].ToString(), Notes = tokens[6].ToString(), Price = tokens[7].ToString() });
                }

            }

            return customerlist;
        }
        #endregion

        #region DoubleClick
        private void lv_DoubleClick(object sender, MouseEventArgs e)
        {
            SelectedValueChanged = true;
            txtName.Text = lstCustomer.SelectedItems[0].SubItems[0].Text;
            txtDescription.Text = lstCustomer.SelectedItems[0].SubItems[1].Text;
            txtStreet.Text = lstCustomer.SelectedItems[0].SubItems[2].Text;
            txtCity.Text = lstCustomer.SelectedItems[0].SubItems[3].Text;
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

            SelectedValueChanged = false;


        }
        #endregion

        #region CustomerDetail
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


            routelist = GetRouteList(filepath);// new List<Routes>();



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

        #endregion

        #region Routes Class
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

            public string notes { get; set; }

            public string price { get; set; }
        }

        #endregion

        #region Save
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
        #endregion

        #region UpdateCustomerTextFile
        private void UpdateCustomerTextFile()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\CustomerLoad.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }


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
        #endregion

        #region UpdateCustomerFromActualList
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
        #endregion

        #region Load Customer Menu
        private void loadCustomersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadCustomers ld = new LoadCustomers();
            ld.Text = "Route Management Software - v 1.0.0.50";
            ld.ShowDialog();

            UpdateCustomerListview();
        }
        #endregion

        #region Print Routes Menu
        private void printRoutesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PrintRoutes pr = new PrintRoutes();
            pr.Text = "Route Management Software - v 1.0.0.50";
            pr.ShowDialog();
        }
        #endregion

        #region Update Routes
        private void updateRoutesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateRoutes uproutes = new UpdateRoutes();
            uproutes.Text = "Route Management Software - v 1.0.0.50";
            uproutes.ShowDialog();
        }
        #endregion

        #region Exit
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #endregion

        #region Add Routes
        private void addRoutesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddRoute ar = new AddRoute();
            ar.Text = "Route Management Software - v 1.0.0.50";
            ar.ShowDialog();

        }
        #endregion

        #region Files Menu
        private void filesLocationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderPath fp = new FolderPath();
            fp.Text = "Route Management Software - v 1.0.0.50";
            fp.ShowDialog();
        }
        #endregion

        #region Add Drivers
        private void addDriversToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddDrivers ad = new AddDrivers();
            ad.Text = "Route Management Software - v 1.0.0.50";
            ad.ShowDialog();
        }
        #endregion

        #region Update Customer Click
        private void btnUpdateCustomer_Click(object sender, EventArgs e)
        {

            if (NameChange)
            {
                int namelist = customerlist.Where(p => p.Name == txtName.Text).Count();

                if (namelist > 0)
                {
                    MessageBox.Show(string.Format("The updated name:'{0}' already exists in the customer list.", txtName.Text), "Duplicate Error", MessageBoxButtons.OK);
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
        #endregion


        #region Update Routelist File
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
        #endregion

        #region Delete
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

        #endregion


        #region Remove Item From Routes
        private void RemoveItemFromRoutes()
        {

            loadedroutelist.RemoveAll(item => item.customername == lstCustomer.SelectedItems[0].SubItems[0].Text && item.description == lstCustomer.SelectedItems[0].SubItems[1].Text
                    && item.street == lstCustomer.SelectedItems[0].SubItems[2].Text && item.city == lstCustomer.SelectedItems[0].SubItems[3].Text
                    && item.state == lstCustomer.SelectedItems[0].SubItems[4].Text && item.zip == lstCustomer.SelectedItems[0].SubItems[5].Text
                    && item.notes == lstCustomer.SelectedItems[0].SubItems[6].Text && item.price == lstCustomer.SelectedItems[0].SubItems[7].Text);


            RepopulateRouteTextFileFromList();
        }

        #endregion

        #region Repopulate Route Text File
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

        #endregion

        #region Repopulate List From ListView
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
        #endregion

        #region Search Text Changed
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
        #endregion

        #region Next
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
        #endregion

        #region Company Name Menu
        private void companyNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CompanyName cn = new CompanyName();
            cn.Text = "Route Management Software - v 1.0.0.50";
            cn.ShowDialog();
        }
        #endregion

        #region Name Changed
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

        #endregion

        #region Description Changed
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
        #endregion

        #region Street Changed
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
        #endregion

        #region City Changed
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

        #endregion

        #region State Changed
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
        #endregion

        #region Zip Changed
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
        #endregion

        #region Price Changed
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
        #endregion

        #region Notes Changed
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

        #endregion


        #region Clear Click
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
        #endregion
    }

}
