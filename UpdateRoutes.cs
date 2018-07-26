using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UpdateCustomers
{
    public partial class UpdateRoutes : Form
    {
        #region Variables
        private List<Driver> driverlist;
        private List<Customer> customerlist;
        private List<Routes> routeList;
        private int currentindexofsearchlist = 0;

        private List<int> searchlist;

        #endregion

        #region Constructor
        public UpdateRoutes()
        {

            InitializeComponent();
            PopulateDriverCombo();
            PopulateDayCombo();
            PopulateFullCustomerList();
            CheckForItems();

            lstFullList.GridLines = true;
            lstNewList.GridLines = true;
            //PopulateRouteList();

            lstFullList.View = View.List;
            lstFullList.View = View.Details;
            lstFullList.FullRowSelect = true;
            lstNewList.View = View.List;
            lstNewList.View = View.Details;
            lstNewList.FullRowSelect = true;
            radAsc.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            radDesc.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);
            radNone.CheckedChanged += new EventHandler(radioButtons_CheckedChanged);

            //  FormBorderStyle = FormBorderStyle.Fixed3D;
            WindowState = FormWindowState.Maximized;

            lblFulLCustomerList.Location = new Point(100, 30);
            lblNewList.Location = new Point(900, 30);

            grpsort.Location = new Point(545, 30);
            btnRightArrow.Location = new Point(550, 130);


            btnSecondMoveUp.Location = new Point(550, 240);
            btnSecondMoveDown.Location = new Point(550, 350);
            btnSaveRoute.Location = new Point(550, 460);
            btnRemove.Location = new Point(550, 570);

            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip1.SetToolTip(btnSaveRoute, "Save Route");
            toolTip1.SetToolTip(btnSecondMoveUp, "Move Item Up In List");
            toolTip1.SetToolTip(btnSecondMoveDown, "Move Item Down In List");
            toolTip1.SetToolTip(btnRightArrow, "Move Selected Item To New List");
            toolTip1.SetToolTip(btnRemove, "Remove Item From List");


            //   btnFirstMoveUp.Location = new Point(400, 600);
            //  btnFirstMoveDown.Location = new Point(400, 625);



            txtSearch.Location = new Point(150, 10);
            lblSearch.Location = new Point(10, 10);

            btnNext.Location = new Point(350, 10);
            //button1.Location = new Point(405, 60);
            lblitemsearch.Location = new Point(445, 10);
            lblsearchamount.Location = new Point(570, 10);



            lstFullList.Font = new Font(lstFullList.Font.FontFamily, 10);
            lstNewList.Font = new Font(lstNewList.Font.FontFamily, 10);

            // ctoDriver.Location = new Point(900, 550);
            lblDriver.Location = new Point(700, 10);
            cboDriver.Location = new Point(800, 10);
            lblDay.Location = new Point(970, 10);
            cboDay.Location = new Point(1070, 10);



            lstFullList.AllowDrop = true;
            cboDay.DropDownStyle = ComboBoxStyle.DropDownList;


            lstNewList.AllowDrop = true;
            //     lstNewList.DragDrop += new DragEventHandler(lv2_DragDrop);
            lstNewList.DragOver += new DragEventHandler(lv2_DragOver);
            //  lstNewList.DragEnter += new DragEventHandler(lstNewList_DragEnter);

            lstFullList.AutoResizeColumn(0,
   ColumnHeaderAutoResizeStyle.ColumnContent);
            lstFullList.AutoResizeColumn(1,
               ColumnHeaderAutoResizeStyle.ColumnContent);
            lstFullList.AutoResizeColumn(2,
               ColumnHeaderAutoResizeStyle.ColumnContent);
        }
        #endregion

        #region radButtons
        private void radioButtons_CheckedChanged(object sender, EventArgs e)
        {
            if (radAsc.Checked)
            {
                lstFullList.Sorting = SortOrder.Ascending;
            }
            if (radDesc.Checked)
            {
                lstFullList.Sorting = SortOrder.Descending;
            }
            if (radNone.Checked)
            {
                lstFullList.Sorting = SortOrder.None;
                lstFullList.Items.Clear();
                PopulateFullCustomerList();
                //lstCustomer.Clear();
                //  lstFullList.Items.Clear();

            }
        }

        #endregion

        #region CheckFor0
        private void CheckForItems()
        {
            if (lstFullList.Items.Count == 0)
            {
                MessageBox.Show("Currently we don't have any items to use for the routes.  Our customer list will be empty until we go back and add some customers..", "Need Items", MessageBoxButtons.OK);
                //Application.Exit();
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



        #region PopulateFullCustomerList
        private void PopulateFullCustomerList()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\CustomerLoad.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }


            customerlist = GetCustomerList(filepath);//new List<Customer>();



            foreach (var item in customerlist)
            {
                ListViewItem lstItem = new ListViewItem(item.Name);
                lstItem.SubItems.Add(item.StreetNumber);
                lstItem.SubItems.Add(item.City);
                lstItem.SubItems.Add(item.State);
                lstItem.SubItems.Add(item.Zip);
                lstItem.SubItems.Add(item.Description);
                lstItem.SubItems.Add(item.Notes);
                lstItem.SubItems.Add(item.Price);

                lstFullList.Items.Add(lstItem);


                //  lstFullList.Items.Add(item.Name + "|" + item.StreetNumber);
            }

        }
        #endregion

        #region GetCustomerList
        private List<Customer> GetCustomerList(string filepath)
        {
            customerlist = new List<Customer>();

            if (File.Exists(filepath))
            {
                var lines = File.ReadAllLines(filepath);
                foreach (var line in lines)
                {

                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = line;
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);
                    //  string encryptstring = cipher.EncryptString(stringtobeencrypted);

                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;


                    string weirdcharcters = "|";
                    string[] tokens = decryptedstring.Split(new[] { weirdcharcters }, StringSplitOptions.None);


                    customerlist.Add(new Customer { Name = tokens[0].ToString(), Description = tokens[1].ToString(), StreetNumber = tokens[2].ToString(), City = tokens[3].ToString(), State = tokens[4].ToString(), Zip = tokens[5].ToString(), Notes = tokens[6].ToString(), Price = tokens[7].ToString() });
                }
            }

            return customerlist;
        }

        #endregion

        #region DayComboBox
        private void PopulateDayCombo()
        {
            cboDay.Items.Add("Monday");
            cboDay.Items.Add("Tuesday");
            cboDay.Items.Add("Wednesday");
            cboDay.Items.Add("Thursday");
            cboDay.Items.Add("Friday");
            cboDay.Items.Add("Saturday");
            cboDay.Items.Add("Sunday");


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

        #region PopulateDriverCombo

        public void PopulateDriverCombo()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Drivers.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            if (File.Exists(filepath))
            {
                driverlist = GetDriverList(filepath);

                foreach (var item in driverlist)
                {
                    cboDriver.Items.Add(item.Name);
                }

            }
            else
            {
                MessageBox.Show("Looks like there aren't any drivers added yet.  We won't have any drivers in the list until we go back to the Driver Form and add them.", "Driver Error", MessageBoxButtons.OK);
            }
        }

        #endregion

        #region Driver Class

        public class Driver
        {
            public string Name { get; set; }
            public string Address { get; set; }
        }
        #endregion

        #region ListView Drags
        private void lv_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                // create array or collection for all selected items
                var items = new List<ListViewItem>();
                // add dragged one first
                items.Add((ListViewItem)e.Item);
                // optionally add the other selected ones
                foreach (ListViewItem lvi in lstFullList.SelectedItems)
                {
                    if (!items.Contains(lvi))
                    {
                        items.Add(lvi);
                    }
                }
                // pass the items to move...
                lstFullList.DoDragDrop(items, DragDropEffects.Move);


                lstNewList.AutoResizeColumn(0,
ColumnHeaderAutoResizeStyle.ColumnContent);
                lstNewList.AutoResizeColumn(1,
                   ColumnHeaderAutoResizeStyle.ColumnContent);
                lstNewList.AutoResizeColumn(2,
                   ColumnHeaderAutoResizeStyle.ColumnContent);

            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }

        private void lv_DragOver(object sender, DragEventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (e.Data.GetDataPresent(typeof(List<ListViewItem>)))
                {
                    e.Effect = DragDropEffects.Move;
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }

        private void lv_DragDrop(object sender, DragEventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (e.Data.GetDataPresent(typeof(List<ListViewItem>)))
                {
                    if (lstNewList.SelectedItems.Count > 0)
                    {

                        ListViewItem item = lstNewList.SelectedItems[0];

                        lstNewList.Items.Remove(item);

                    }

                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }

        private void lv2_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                // create array or collection for all selected items
                var items = new List<ListViewItem>();
                // add dragged one first
                items.Add((ListViewItem)e.Item);
                // optionally add the other selected ones
                foreach (ListViewItem lvi in lstNewList.SelectedItems)
                {
                    if (!items.Contains(lvi))
                    {
                        items.Add(lvi);
                    }
                }
                // pass the items to move...
                lstNewList.DoDragDrop(items, DragDropEffects.Move);
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }

        // this SHOULD look at KeyState to disallow actions not supported
        private void lv2_DragOver(object sender, DragEventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (e.Data.GetDataPresent(typeof(List<ListViewItem>)))
                {
                    e.Effect = DragDropEffects.Move;
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }

        private void lv2_DragDrop(object sender, DragEventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (e.Data.GetDataPresent(typeof(List<ListViewItem>)))
                {
                    if (lstFullList.SelectedItems.Count > 0)
                    {
                        ListViewItem item = lstFullList.SelectedItems[0];

                        //foreach (ListViewItem lvi in items)
                        //{
                        // LVI obj can only belong to one LVI, remove

                       // lstFullList.Items.Remove(item);
                        lstNewList.Items.Add((ListViewItem)item.Clone());
                        // lstNewList.Items.Add(lvi);
                    }
                    //var items = (List<ListViewItem>)e.Data.GetData(typeof(List<ListViewItem>));
                    // move to dest LV
                    lstNewList.AutoResizeColumn(0,
ColumnHeaderAutoResizeStyle.ColumnContent);
                    lstNewList.AutoResizeColumn(1,
                       ColumnHeaderAutoResizeStyle.ColumnContent);
                    lstNewList.AutoResizeColumn(2,
                       ColumnHeaderAutoResizeStyle.ColumnContent);
                    //}
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }
        #endregion

        #region save

        private void btnSaveRoute_Click(object sender, EventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (cboDriver.Text == "" || cboDay.Text == "")
                {
                    MessageBox.Show("The Driver or Day isn't filled out", "Additional Info Needed", MessageBoxButtons.OK);
                    return;
                }


                AddUpdatedListViewToRouteList();

                MessageBox.Show("Route Saved", "Update", MessageBoxButtons.OK);
            }
            else
            {
                MessageBox.Show("Driver or Day are missing to save route.", "Missing Information", MessageBoxButtons.OK);
            }

            lstNewList.AutoResizeColumn(0,
               ColumnHeaderAutoResizeStyle.ColumnContent);
            lstNewList.AutoResizeColumn(1,
               ColumnHeaderAutoResizeStyle.ColumnContent);
            lstNewList.AutoResizeColumn(2,
               ColumnHeaderAutoResizeStyle.ColumnContent);

        }
        #endregion

        #region SelectionChanged

        private void cboDay_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstNewList.Items.Clear();


            if (cboDriver.Text != "")
            {
               
                string path = (string)Properties.Settings.Default["Path"];
                path += "\\RouteManagementFiles";

                string filepath = path + "\\Routes.txt";
                
                routeList = GetRouteList(filepath); 
                
                var actualroute = routeList.Where(p => p.driver == cboDriver.Text && p.day == cboDay.Text);
                
                foreach (var item in actualroute)
                {
                    ListViewItem lstItem = new ListViewItem(item.customername);
                    lstItem.SubItems.Add(item.street);
                    lstItem.SubItems.Add(item.city);
                    lstItem.SubItems.Add(item.state);
                    lstItem.SubItems.Add(item.zip);
                    lstItem.SubItems.Add(item.description);
                    lstItem.SubItems.Add(item.notes);
                    lstItem.SubItems.Add(item.price);
                    lstNewList.Items.Add(lstItem);
                }




            }

            lstNewList.AutoResizeColumn(0,
   ColumnHeaderAutoResizeStyle.ColumnContent);
            lstNewList.AutoResizeColumn(1,
               ColumnHeaderAutoResizeStyle.ColumnContent);
            lstNewList.AutoResizeColumn(2,
               ColumnHeaderAutoResizeStyle.ColumnContent);
        }

        #endregion

        #region GetRouteList
        public List<Routes> GetRouteList(string filepath)
        {
            string line;
            routeList = new List<Routes>();
            using (FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            using (StreamReader sr = new StreamReader(fs))
            {
                while ((line = sr.ReadLine()) != null)
                // foreach (var line in sr.ReadToEnd())
                {

                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = (line.ToString());
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);


                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;


                    string[] tokens = decryptedstring.Split('|');
                    routeList.Add(new Routes { driver = tokens[0].ToString(), day = tokens[1].ToString(), customername = tokens[2].ToString(), street = tokens[3].ToString(), city = tokens[4].ToString(), state = tokens[5].ToString(), zip = tokens[6].ToString(), description = tokens[7].ToString(), notes = tokens[8].ToString(), price = tokens[9].ToString() });
                }
            }

            return routeList;

        }
        #endregion


        #region MoveupClick
        private void btnSecondMoveUp_Click(object sender, EventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                foreach (ListViewItem lvi in lstNewList.SelectedItems)
                {
                    int index = lvi.Index - 1;


                    if (lstNewList.Items.Count - 1 >= index && lstNewList.Items.Count > 1 && index > -1)
                    {
                        lstNewList.Items.RemoveAt(lvi.Index);
                        lstNewList.Items.Insert(index, lvi);


                    }
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }
        #endregion


        #region MoveDownClick
        private void btnSecondMoveDown_Click(object sender, EventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                foreach (ListViewItem lvi in lstNewList.SelectedItems)
                {
                    int index = lvi.Index + 1;

                    if (lstNewList.Items.Count - 1 >= index && lstNewList.Items.Count > 1)
                    {
                        lstNewList.Items.RemoveAt(lvi.Index);
                        lstNewList.Items.Insert(index, lvi);

                    }
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }
        #endregion

        #region TextChanged
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            currentindexofsearchlist = 0;
            int values = 0;
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                foreach (ListViewItem item in lstFullList.Items)
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
                    lblsearchamount.Text = values.ToString();
                   
                }

            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }


            searchlist = new List<int>();
            

            foreach(ListViewItem item in lstFullList.Items)
            {
                if(item.BackColor == Color.CornflowerBlue)
                searchlist.Add(item.Index);
            }
            if (searchlist.Count > 0)
            lstFullList.TopItem = lstFullList.Items[searchlist.FirstOrDefault()];
        }

        #endregion


        #region UpdateRouteFile
        private void UpdateRouteListFile()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            using (FileStream f = new FileStream(filepath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (StreamWriter s = new StreamWriter(f))
            {
                foreach (var item in routeList)
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

        #region Remove
        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (lstNewList.SelectedItems.Count > 0)
                {
                    lstNewList.SelectedItems[0].Remove();

                    routeList.RemoveAll(r => r.driver == cboDriver.Text && r.day == cboDay.Text);


                    AddUpdatedListViewToRouteList();

                }
                else
                {
                    MessageBox.Show("No Items Selected.", "Select Item", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
            
        }
        #endregion

        #region ListViewToRouteList
        private void AddUpdatedListViewToRouteList()
        {
            routeList.RemoveAll(r => r.driver == cboDriver.Text && r.day == cboDay.Text);

            foreach (ListViewItem itemRow in this.lstNewList.Items)
            {
                routeList.Add(new Routes { driver = cboDriver.Text, day = cboDay.Text, customername = itemRow.SubItems[0].Text, street = itemRow.SubItems[1].Text, city = itemRow.SubItems[2].Text, state = itemRow.SubItems[3].Text, zip = itemRow.SubItems[4].Text, description = itemRow.SubItems[5].Text, notes = itemRow.SubItems[6].Text, price = itemRow.SubItems[7].Text });
            }

            UpdateRouteListFile();
        }
        #endregion


        #region Next

        private void btnNext_Click(object sender, EventArgs e)
        {

            if (lstFullList.SelectedItems.Count > 0)
            {
                if (searchlist.Count > 0)
                {
                    if (currentindexofsearchlist > searchlist.Count - 1)
                        currentindexofsearchlist = 0;


                    lstFullList.TopItem = lstFullList.Items[searchlist[currentindexofsearchlist]];
                }
            }

            currentindexofsearchlist++;
        }
        #endregion


        #region RightArrow
        private void btnRightArrow_Click(object sender, EventArgs e)
        {
            if (cboDay.Text != "" && cboDriver.Text != "")
            {
                if (lstFullList.SelectedItems.Count > 0)
                {

                    ListViewItem item = lstFullList.SelectedItems[0];

                    lstNewList.Items.Add((ListViewItem)item.Clone());


                    lstNewList.AutoResizeColumn(0,
       ColumnHeaderAutoResizeStyle.ColumnContent);
                    lstNewList.AutoResizeColumn(1,
                       ColumnHeaderAutoResizeStyle.ColumnContent);
                    lstNewList.AutoResizeColumn(2,
                       ColumnHeaderAutoResizeStyle.ColumnContent);

                }
                else
                {
                    MessageBox.Show("No Items selected.", "Select Items", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Please choose a driver and day to begin.", "Needing Input", MessageBoxButtons.OK);
            }
        }
    }
    #endregion
}

