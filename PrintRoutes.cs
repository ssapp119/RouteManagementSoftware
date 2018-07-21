using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Linq;

namespace UpdateCustomers
{
    public partial class PrintRoutes : Form
    {
        private int value = 0;
        //  private System.ComponentModel.Container components;
        private List<RoutePrinterList> pRouteList;
        private List<string> cbovalues;
        private int counter = 3;
        private int i = 0;
        private Dictionary<string, string> reportvalues;
        private Dictionary<string, string> categoryvalues;
        private List<string> itemsforreport;
        private PrintDocument recordDoc;
        private string driverday;
        private string category;
        private static int numberofitems = 0;
        private static int numberofitemsprintedsofar = 0;
        private static float y1 = 98.0F;
        private bool NextPage = false;
        private bool PrintAdditionalItems = false;
        private SortedDictionary<int, string> additionalitems;
        private int Index = 0;
        private int checkboxcount = 0;
        public int fontsize = 0;
        public int fontsizeCategory = 18;


        public PrintRoutes()
        {
            InitializeComponent();
            PopulateDriverCombo();


            cboBusinessName.Enabled = true;
            cboDescription.Enabled = false;
            cboStreet.Enabled = true;
            cboCity.Enabled = false;
            cboState.Enabled = false;
            cboZip.Enabled = false;
            cboNotes.Enabled = true;
            cboPrice.Enabled = false;

            cboBusinessName.Text = "1";
            cboStreet.Text = "2";
            cboNotes.Text = "3";

            //   chkBusinessName.Checked = true;
            //    chkStreet.Checked = true;
            //    chkNotes.Checked = true;



            for (i = 1; i <= counter; i++)
            {
                cboBusinessName.Items.Add(i.ToString());
                cboStreet.Items.Add(i.ToString());
                cboNotes.Items.Add(i.ToString());
            }



            //if (chkBusinessName.Checked == true)
            //{
            //    cboBusinessName.Enabled = true;
            //}



            //  cboDay.SelectedIndex = -1;
        }

        private void PopulateDriverCombo()
        {
            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }



            pRouteList = new List<RoutePrinterList>();

            if (File.Exists(filepath))
            {

                var lines = File.ReadAllLines(filepath);
                foreach (var line in lines)
                {
                    StringCipher cipher = new StringCipher();
                    //string userid = txtUserID.Text.ToString();
                    string stringtobedecrypted = line;
                    string decryptedstring = cipher.Decrypt(stringtobedecrypted);

                    if (string.IsNullOrEmpty(decryptedstring))
                        continue;




                    string weirdcharcters = "|";
                    string[] tokens = decryptedstring.Split(new[] { weirdcharcters }, StringSplitOptions.None);


                    pRouteList.Add(new RoutePrinterList { drivername = tokens[0].ToString(), day = tokens[1].ToString(), businessname = tokens[2].ToString(), street = tokens[3].ToString() });

                }




                var distinctdrivers = pRouteList
      .GroupBy(c => c.drivername)
      .Select(g => g.First());

                foreach (var item in distinctdrivers)
                {
                    cboDrivers.Items.Add(item.drivername);
                }


            }
        }

        public void PrintReceiptPage(object sender, PrintPageEventArgs e)
        {


            value = 0;

            switch (cbovalues.Count)
            {
                case 1:
                    category = categoryvalues["1"];

                    break;
                case 2:
                    category += categoryvalues["1"];
                    category += categoryvalues["2"];
                    break;
                case 3:
                    category = categoryvalues["1"].PadRight(50);
                    category += categoryvalues["2"].PadRight(50);
                    category += categoryvalues["3"].PadRight(50);
                    break;

                case 4:
                    category = categoryvalues["1"];
                    category += categoryvalues["2"];
                    category += categoryvalues["3"];
                    category += categoryvalues["4"];
                    break;

                case 5:
                    category = categoryvalues["1"];
                    category += categoryvalues["2"];
                    category += categoryvalues["3"];
                    category += categoryvalues["4"];
                    category += categoryvalues["5"];
                    break;

                case 6:

                    category += categoryvalues["1"];
                    category += categoryvalues["2"];
                    category += categoryvalues["3"];
                    category += categoryvalues["4"];
                    category += categoryvalues["5"];
                    category += categoryvalues["6"];
                    break;

                case 7:
                    category = categoryvalues["1"];
                    category += categoryvalues["2"];
                    category += categoryvalues["3"];
                    category += categoryvalues["4"];
                    category += categoryvalues["5"];
                    category += categoryvalues["6"];
                    category += categoryvalues["7"];
                    break;

                case 8:
                    category = categoryvalues["1"];
                    category += categoryvalues["2"];
                    category += categoryvalues["3"];
                    category += categoryvalues["4"];
                    category += categoryvalues["5"];
                    category += categoryvalues["6"];
                    category += categoryvalues["7"];
                    category += categoryvalues["8"];
                    break;
            }






            itemsforreport = new List<string>();

            float x = 10.0F;
            float y = 10.0F;
            //  float y1 = 98;
            float width = 1000.0F;
            // float width = 270.0F; // max width I found through trial and error
            float height = 0F;


            Font drawFontArialBold = new Font("Arial", fontsize, FontStyle.Bold);
            Font drawFontCategoryBold = new Font("Arial", 18, FontStyle.Bold);
            //    Font drawFontArial16Bold = new Font("Arial", 20, FontStyle.Bold);
            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Set format of string.
            StringFormat drawFormatCenter = new StringFormat();
            drawFormatCenter.Alignment = StringAlignment.Center;
            StringFormat drawFormatLeft = new StringFormat();
            drawFormatLeft.Alignment = StringAlignment.Near;
            StringFormat drawFormatRight = new StringFormat();
            drawFormatRight.Alignment = StringAlignment.Far;

            // Draw string to screen.
            //  string text = "Can Redemption Center";
            string companyname = Properties.Settings.Default.CompanyName;
            e.Graphics.DrawString(companyname, drawFontCategoryBold, drawBrush, new RectangleF(x, y, width, height), drawFormatCenter);
            y += e.Graphics.MeasureString(companyname, drawFontCategoryBold).Height;

            driverday = "           Driver: " + cboDrivers.Text + "                                   Day: " + cboDay.Text;

            e.Graphics.DrawString(driverday, drawFontCategoryBold, drawBrush, new RectangleF(x, y, width, height), drawFormatLeft);
            y += e.Graphics.MeasureString(driverday, drawFontCategoryBold).Height;

            // text = "Carroll, IA";

            e.Graphics.DrawString(" ", drawFontCategoryBold, drawBrush, new RectangleF(x, y, width, height), drawFormatCenter);
            y += e.Graphics.MeasureString(" ", drawFontCategoryBold).Height;


            //   int x1 = (int)x;
            //   int y1 = (int)y;
            Pen blackPen = new Pen(Color.Black, 2);



            //if (categoryvalues.Count == 1)
            //{
            //    e.Graphics.DrawString(categoryvalues["1"], drawFontArialBold, drawBrush, new RectangleF(x, y, width, height), drawFormatCenter);
            //    y += e.Graphics.MeasureString(category, drawFontArialBold).Height;
            //}
            //else if (categoryvalues.Count == 2)
            //{
            //    e.Graphics.DrawString(categoryvalues["1"], drawFontArialBold, drawBrush, new RectangleF(x + 200, y, width, height), drawFormatLeft);
            //    e.Graphics.DrawString(categoryvalues["2"], drawFontArialBold, drawBrush, new RectangleF(x + 700, y, width, height), drawFormatLeft);
            //    y += e.Graphics.MeasureString(category, drawFontArialBold).Height;
            //}
            //else if (categoryvalues.Count == 3)
            //{

            //    int xbob1 = (int)x;
            //    int ybob1 = (int)y;

            //    int xbob2 = (int)x + 350;
            //    int ybob2 = (int)y;

            //    int xbob3 = (int)x + 700;
            //    int ybob3 = (int)y;
            //    // Create rectangle.

            //    Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 57);
            //    e.Graphics.DrawRectangle(blackPen, bob1rect);

            //    Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
            //    e.Graphics.DrawRectangle(blackPen, bob1rect2);

            //    Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
            //    e.Graphics.DrawRectangle(blackPen, bob1rect3);




            //    e.Graphics.DrawString(categoryvalues["1"], drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
            //    e.Graphics.DrawString(categoryvalues["2"], drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
            //    e.Graphics.DrawString(categoryvalues["3"], drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
            //    y += e.Graphics.MeasureString(categoryvalues["1"], drawFontArialBold).Height;
            //}

            //foreach (ListViewItem itemRow in this.lstViewPOS.Items)
            //{

            string path = (string)Properties.Settings.Default["Path"];
            path += "\\RouteManagementFiles";

            string filepath = path + "\\Routes.txt";



            pRouteList = new List<RoutePrinterList>();

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


                    pRouteList.Add(new RoutePrinterList { drivername = tokens[0].ToString(), day = tokens[1].ToString(), businessname = tokens[2].ToString(), street = tokens[3].ToString(), city = tokens[4].ToString(), state = tokens[5].ToString(), zip = tokens[6].ToString(), description = tokens[7].ToString(), notes = tokens[8].ToString(), price = tokens[9].ToString() });

                }

            }
            var actualroute = pRouteList.Where(p => p.drivername == cboDrivers.Text && p.day == cboDay.Text);
            int purchCount = actualroute.Count();

            foreach (var item in actualroute.Skip(Index))
            {
                reportvalues.Clear();
                // categoryvalues.Clear();

                //  cboDriver.Items.Add(item.Name);
                string bob1 = string.Empty;
                string bob2 = string.Empty;
                string bob3 = string.Empty;
                string bob4 = string.Empty;
                string bob5 = string.Empty;
                string bob6 = string.Empty;
                string bob7 = string.Empty;
                string bob8 = string.Empty;

                string[] bob1substrings = new string[3];
                string[] bob2substrings = new string[3];
                string[] bob3substrings = new string[3];






                if (cbovalues.Count == 1)
                {
                    if (chkBusinessName.Checked == true)
                    {

                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {

                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }

                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);


                    }

                    if (chkPrice.Checked == true)
                    {
                        reportvalues.Add(cboPrice.Text, item.price);


                    }
                }
                else if (cbovalues.Count == 2)
                {
                    if (chkBusinessName.Checked == true)
                    {
                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {
                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }

                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);


                    }

                    if (chkPrice.Checked == true)
                    {
                        reportvalues.Add(cboPrice.Text, item.price);


                    }




                }
                else if (cbovalues.Count == 3)
                {
                    if (chkBusinessName.Checked == true)
                    {
                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {
                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }

                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);


                    }

                    if (chkPrice.Checked == true)
                    {
                        reportvalues.Add(cboPrice.Text, item.price);


                    }
                }
                else if (cbovalues.Count == 4)
                {
                    if (chkBusinessName.Checked == true)
                    {
                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {
                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }
                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);


                    }

                    if (chkPrice.Checked == true)
                    {
                        reportvalues.Add(cboPrice.Text, item.price);

                    }
                }
                else if (cbovalues.Count == 5)
                {
                    if (chkBusinessName.Checked == true)
                    {
                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {
                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }
                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);


                    }

                    if (chkPrice.Checked == true)
                    {
                        reportvalues.Add(cboPrice.Text, item.price);


                    }
                }
                else if (cbovalues.Count == 6)
                {
                    if (chkBusinessName.Checked == true)
                    {
                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {
                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }
                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);

                    }

                    if (chkPrice.Checked == true)
                    {
                        reportvalues.Add(cboPrice.Text, item.price);


                    }
                }
                else if (cbovalues.Count == 7)
                {
                    if (chkBusinessName.Checked == true)
                    {
                        reportvalues.Add(cboBusinessName.Text, item.businessname);

                    }
                    if (chkDescription.Checked == true)
                    {
                        reportvalues.Add(cboDescription.Text, item.description);

                    }
                    if (chkStreet.Checked == true)
                    {
                        reportvalues.Add(cboStreet.Text, item.street);

                    }
                    if (chkCity.Checked == true)
                    {
                        reportvalues.Add(cboCity.Text, item.city);

                    }
                    if (chkState.Checked == true)
                    {
                        reportvalues.Add(cboState.Text, item.state);

                    }
                    if (chkZip.Checked == true)
                    {
                        reportvalues.Add(cboZip.Text, item.zip);


                    }
                    if (chkNotes.Checked == true)
                    {
                        reportvalues.Add(cboNotes.Text, item.notes);


                    }

                    if (chkPrice.Checked == true)
                    {

                        reportvalues.Add(cboPrice.Text, item.price);


                    }
                }


                switch (cbovalues.Count)
                {
                    case 1:

                        bob1 = reportvalues["1"].PadRight(80);
                        if (bob1.Length > 40)
                        {
                            bob1substrings[0] = bob1.Substring(0, 20);
                            bob1substrings[1] = bob1.Substring(21, 20);
                            bob1substrings[2] = bob1.Substring(40, 20);
                        }


                        break;

                    case 2:
                        bob1 = reportvalues["1"].PadRight(35);
                        bob2 += reportvalues["2"].PadRight(35);

                        break;

                    case 3:
                        bob1 = reportvalues["1"].PadRight(60);
                        bob2 = reportvalues["2"].PadRight(60);
                        bob3 = reportvalues["3"].PadRight(60);

                        if (bob1.Length > 40)
                        {
                            bob1substrings[1] = bob1.Substring(0, 20);
                            bob1substrings[2] = bob1.Substring(21, 20);
                            //            bob1substrings[3] = bob1.Substring(40, 20);
                        }

                        if (bob2.Length > 40)
                        {
                            string bob2substring = bob2.Substring(0, 20);
                            string bob2substring2 = bob2.Substring(21, 20);
                        }

                        if (bob3.Length > 40)
                        {
                            string bob1substring = bob1.Substring(0, 20);
                            string bob1substring2 = bob1.Substring(21, 20);
                        }



                        break;

                    case 4:
                        bob1 = reportvalues["1"];
                        bob2 += reportvalues["2"];
                        bob3 += reportvalues["3"];
                        bob4 += reportvalues["4"];


                        break;

                    case 5:
                        bob1 = reportvalues["1"];
                        bob2 += reportvalues["2"];
                        bob3 += reportvalues["3"];
                        bob4 += reportvalues["4"];
                        bob5 += reportvalues["5"];


                        break;

                    case 6:
                        bob1 = reportvalues["1"];
                        bob2 += reportvalues["2"];
                        bob3 += reportvalues["3"];
                        bob4 += reportvalues["4"];
                        bob5 += reportvalues["5"];
                        bob6 += reportvalues["6"];


                        break;

                    case 7:
                        bob1 = reportvalues["1"];
                        bob2 += reportvalues["2"];
                        bob3 += reportvalues["3"];
                        bob4 += reportvalues["4"];
                        bob5 += reportvalues["5"];
                        bob6 += reportvalues["6"];
                        bob7 += reportvalues["7"];

                        break;

                    case 8:
                        bob1 = reportvalues["1"];
                        bob2 += reportvalues["2"];
                        bob3 += reportvalues["3"];
                        bob4 += reportvalues["4"];
                        bob5 += reportvalues["5"];
                        bob6 += reportvalues["6"];
                        bob7 += reportvalues["7"];
                        bob8 += reportvalues["8"];

                        break;

                }

                if (value == 0)
                {
                    value++;
                    if (categoryvalues.Count == 1)
                    {
                        bool adjustedvalues = false;



                        //check if state, zip, 

                        int xbob1 = (int)x;
                        int ybob1 = (int)y1;


                        Rectangle bob1rect;
                        // Create rectangle.
                        if (radbtn10.Checked)
                        {
                            bob1rect = new Rectangle(xbob1, ybob1, 1050, 57);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn15.Checked)
                        {
                            bob1rect = new Rectangle(xbob1, ybob1, 1050, 28);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);
                            y += bob1rect.Height - 5;
                        }
                        if (radbtn20.Checked)
                        {
                            bob1rect = new Rectangle(xbob1, ybob1, 1050, 28);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);
                            y += bob1rect.Height - 5;
                        }


                        if (radbtn25.Checked)
                        {
                            bob1rect = new Rectangle(xbob1, ybob1, 1050, 28);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);
                            y += bob1rect.Height - 5;
                        }
                        // bob1rect = new Rectangle(xbob1, ybob1, 1050, 57);
                        //Rectangle bob1rect = new Rectangle(xbob1, ybob1, 1050, 57);
                        //   e.Graphics.DrawRectangle(blackPen, bob1rect);

                        //Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
                        ////   e.Graphics.DrawRectangle(blackPen, bob1rect2);


                        //Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                        ////   e.Graphics.DrawRectangle(blackPen, bob1rect3);






                        //e.Graphics.DrawString(categoryvalues["2"], drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                        //e.Graphics.DrawRectangle(blackPen, bob1rect2);

                        ////               e.Graphics.DrawString(categoryvalues["3"], drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                        //               e.Graphics.DrawRectangle(blackPen, bob1rect3);

                        //  y += e.Graphics.MeasureString(categoryvalues["1"], drawFontArialBold).Height;
                     //   y += 32;
                    }
                    else if (categoryvalues.Count == 2)
                    {
                        bool adjustedvalues = false;



                        //check if state, zip, 

                        int xbob1 = (int)x;
                        int ybob1 = (int)y1;

                        int xbob2 = (int)x + 525;
                        int ybob2 = (int)y1;

                        //int xbob3 = (int)x + 700;
                        //int ybob3 = (int)y1;
                        // Create rectangle.

                        if (radbtn10.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 57);
                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 57);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);



                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);
                            y += bob1rect.Height - 5;

                        }

                        if (radbtn15.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 28);
                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 28);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);



                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn20.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 28);
                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 28);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);



                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn25.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 28);
                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 28);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);



                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            y += bob1rect.Height - 5;

                        }
                        //   e.Graphics.DrawRectangle(blackPen, bob1rect);


                        //   e.Graphics.DrawRectangle(blackPen, bob1rect2);


                        //Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                        ////   e.Graphics.DrawRectangle(blackPen, bob1rect3);




                        ////               e.Graphics.DrawString(categoryvalues["3"], drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                        //               e.Graphics.DrawRectangle(blackPen, bob1rect3);

                        //   y += e.Graphics.MeasureString(categoryvalues["1"], drawFontArialBold).Height;


                    }
                    else if (categoryvalues.Count == 3)
                    {
                        //  bool adjustedvalues = false;



                        //check if state, zip, 

                        int xbob1 = (int)x;
                        int ybob1 = (int)y1;

                        int xbob2 = (int)x + 350;
                        int ybob2 = (int)y1;

                        int xbob3 = (int)x + 700;
                        int ybob3 = (int)y1;
                        // Create rectangle.


                        if (radbtn10.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 57);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect2);


                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            y += bob1rect.Height - 5;
                        }

                        if (radbtn15.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect2);


                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            y += bob1rect.Height - 5;

                        }

                        if (radbtn20.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect2);


                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            y += bob1rect.Height - 5;

                        }

                        if (radbtn25.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect2);


                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 28);
                            //   e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            y += bob1rect.Height - 5;

                        }



                        // y += e.Graphics.MeasureString(categoryvalues["1"], drawFontArialBold).Height;
                    }



                    else if (categoryvalues.Count == 4)
                    {



                        int xbob1 = (int)x;
                        int ybob1 = (int)y1;

                        int xbob2 = (int)x + 262;
                        int ybob2 = (int)y1;

                        int xbob3 = (int)x + 524;
                        int ybob3 = (int)y1;

                        int xbob4 = (int)x + 787;
                        int ybob4 = (int)y1;
                        // Create rectangle.

                        if (radbtn10.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn15.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn20.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn25.Checked)
                        {

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }
                     //   y += e.Graphics.MeasureString(categoryvalues["1"], drawFontCategoryBold).Height;
                    }

                    else if (categoryvalues.Count == 5)
                    {

                        int xbob1 = (int)x;
                        int ybob1 = (int)y1;

                        int xbob2 = (int)x + 210;
                        int ybob2 = (int)y1;

                        int xbob3 = (int)x + 420;
                        int ybob3 = (int)y1;

                        int xbob4 = (int)x + 630;
                        int ybob4 = (int)y1;

                        int xbob5 = (int)x + 840;
                        int ybob5 = (int)y1;
                        // Create rectangle.


                        if (radbtn10.Checked)
                        {
                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect5);




                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["5"], drawFontCategoryBold, drawBrush, bob1rect5, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn15.Checked)
                        {
                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 28);
                            e.Graphics.DrawRectangle(blackPen, bob1rect5);




                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["5"], drawFontCategoryBold, drawBrush, bob1rect5, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn20.Checked)
                        {
                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect5);




                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["5"], drawFontCategoryBold, drawBrush, bob1rect5, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }

                        if (radbtn25.Checked)
                        {
                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 57);
                            e.Graphics.DrawRectangle(blackPen, bob1rect5);




                            e.Graphics.DrawString(categoryvalues["1"], drawFontCategoryBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["2"], drawFontCategoryBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["3"], drawFontCategoryBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["4"], drawFontCategoryBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(categoryvalues["5"], drawFontCategoryBold, drawBrush, bob1rect5, drawFormatCenter);
                            y += bob1rect.Height - 5;
                        }
                      //  y += e.Graphics.MeasureString(categoryvalues["1"], drawFontCategoryBold).Height;
                    }


                    //else if (categoryvalues.Count == 6)
                    //{

                    //    int xbob1 = (int)x;
                    //    int ybob1 = (int)y1;

                    //    int xbob2 = (int)x + 350;
                    //    int ybob2 = (int)y1;

                    //    int xbob3 = (int)x + 700;
                    //    int ybob3 = (int)y1;
                    //    // Create rectangle.

                    //    Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 57);
                    //    e.Graphics.DrawRectangle(blackPen, bob1rect);

                    //    Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
                    //    e.Graphics.DrawRectangle(blackPen, bob1rect2);

                    //    Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                    //    e.Graphics.DrawRectangle(blackPen, bob1rect3);




                    //    e.Graphics.DrawString(categoryvalues["1"], drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                    //    e.Graphics.DrawString(categoryvalues["2"], drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                    //    e.Graphics.DrawString(categoryvalues["3"], drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                    //    y += e.Graphics.MeasureString(categoryvalues["1"], drawFontArialBold).Height;
                    //}


                    //else if (categoryvalues.Count == 7)
                    //{

                    //    int xbob1 = (int)x;
                    //    int ybob1 = (int)y1;

                    //    int xbob2 = (int)x + 350;
                    //    int ybob2 = (int)y1;

                    //    int xbob3 = (int)x + 700;
                    //    int ybob3 = (int)y1;
                    //    // Create rectangle.

                    //    Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 57);
                    //    e.Graphics.DrawRectangle(blackPen, bob1rect);

                    //    Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
                    //    e.Graphics.DrawRectangle(blackPen, bob1rect2);

                    //    Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                    //    e.Graphics.DrawRectangle(blackPen, bob1rect3);




                    //    e.Graphics.DrawString(categoryvalues["1"], drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                    //    e.Graphics.DrawString(categoryvalues["2"], drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                    //    e.Graphics.DrawString(categoryvalues["3"], drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                    //    y += e.Graphics.MeasureString(categoryvalues["1"], drawFontArialBold).Height;
                    //}
                }

                //    string bob = item.businessname + "      " + item.street + "      ";
                // Do something useful here !
                // e.g 'itemRow.SubItems[count]' <-- Should give you direct access to
                // the item located at co-ordinates(0,0). Once you got it, do something 
                // with it.

                //   text = "Carroll, IA";

                //e.Graphics.DrawString(" ", drawFontArial10Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatCenter);
                //if (radbtn10.Checked)
                //{
                //    y += e.Graphics.MeasureString(" ", drawFontArial10Regular).Height;

                //    //e.Graphics.DrawString(" ", drawFontArial10Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatCenter);
                //    y += e.Graphics.MeasureString(" ", drawFontArial10Regular).Height;
                //}

                Pen blackPen1 = new Pen(Color.Black, 2);



                //int x1 = (int)x;
                //int y1 = (int)y;
                // Create rectangle.
                //   Rectangle rect = new Rectangle(x1, y1, 200, 100);

                // Draw rectangle to screen.
                //      e.Graphics.DrawRectangle(blackPen, rect);



                if (categoryvalues.Count == 1)
                {
                    if (radbtn10.Checked)
                    {

                        if (numberofitems < 10)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 1050, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            //y += 28; //e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }


                    if (radbtn15.Checked)
                    {

                        if (numberofitems < 15)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 1050, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            //   y += 28; //e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }







                    if (radbtn20.Checked)
                    {

                        if (numberofitems < 20)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 1050, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            //  y += e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;
                            //  y += 28;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }



                    if (radbtn25.Checked)
                    {

                        if (numberofitems < 25)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 1050, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            //  y += e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;
                            //    y += 28;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }






                }
                else if (categoryvalues.Count == 2)
                {

                    if (radbtn10.Checked)
                    {
                        if (numberofitems < 10)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 525;
                            int ybob2 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            //  y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn15.Checked)
                    {
                        if (numberofitems < 15)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 525;
                            int ybob2 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            //  y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                    if (radbtn20.Checked)
                    {
                        if (numberofitems < 20)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 525;
                            int ybob2 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            //  y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn25.Checked)
                    {
                        if (numberofitems < 25)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 525;
                            int ybob2 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 525, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 525, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            //  y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                }
                else if (categoryvalues.Count == 3)
                {
                    if (radbtn10.Checked)
                    {
                        if (numberofitems < 10)
                        {
                            //if (numberofitemsprintedsofar <= purchCount)
                            //{
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 350;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 700;
                            int ybob3 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }


                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn15.Checked)
                    {
                        if (numberofitems < 15)
                        {
                            //if (numberofitemsprintedsofar <= purchCount)
                            //{
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 350;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 700;
                            int ybob3 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }


                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn20.Checked)
                    {
                        if (numberofitems < 20)
                        {
                            //if (numberofitemsprintedsofar <= purchCount)
                            //{
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 350;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 700;
                            int ybob3 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }


                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn25.Checked)
                    {
                        if (numberofitems < 25)
                        {
                            //if (numberofitemsprintedsofar <= purchCount)
                            //{
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 350;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 700;
                            int ybob3 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 350, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 350, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 350, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }


                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                }



                else if (categoryvalues.Count == 4)
                {
                    if (radbtn10.Checked)
                    {
                        if (numberofitems < 10)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 262;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 524;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 786;
                            int ybob4 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;

                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                    if (radbtn15.Checked)
                    {
                        if (numberofitems < 15)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 262;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 524;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 786;
                            int ybob4 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;

                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                    if (radbtn20.Checked)
                    {
                        if (numberofitems < 20)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 262;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 524;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 786;
                            int ybob4 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;

                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn25.Checked)
                    {
                        if (numberofitems < 25)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 262;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 524;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 786;
                            int ybob4 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 262, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            //   y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;

                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                }
                else if (categoryvalues.Count == 5)
                {
                    if (radbtn10.Checked)
                    {
                        if (numberofitems < 10)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 210;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 420;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 630;
                            int ybob4 = (int)y;

                            int xbob5 = (int)x + 840;
                            int ybob5 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect5);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(bob5, drawFontArialBold, drawBrush, bob1rect5, drawFormatCenter);
                            //     y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                    if (radbtn15.Checked)
                    {
                        if (numberofitems < 15)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 210;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 420;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 630;
                            int ybob4 = (int)y;

                            int xbob5 = (int)x + 840;
                            int ybob5 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 57);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect5);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(bob5, drawFontArialBold, drawBrush, bob1rect5, drawFormatCenter);
                            //     y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn20.Checked)
                    {
                        if (numberofitems < 20)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 210;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 420;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 630;
                            int ybob4 = (int)y;

                            int xbob5 = (int)x + 840;
                            int ybob5 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect5);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(bob5, drawFontArialBold, drawBrush, bob1rect5, drawFormatCenter);
                            //     y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }

                    if (radbtn25.Checked)
                    {
                        if (numberofitems < 25)
                        {
                            Index++;
                            numberofitems++;
                            int xbob1 = (int)x;
                            int ybob1 = (int)y;

                            int xbob2 = (int)x + 210;
                            int ybob2 = (int)y;

                            int xbob3 = (int)x + 420;
                            int ybob3 = (int)y;

                            int xbob4 = (int)x + 630;
                            int ybob4 = (int)y;

                            int xbob5 = (int)x + 840;
                            int ybob5 = (int)y;
                            // Create rectangle.

                            Rectangle bob1rect = new Rectangle(xbob1, ybob1, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect);

                            Rectangle bob1rect2 = new Rectangle(xbob2, ybob2, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect2);

                            Rectangle bob1rect3 = new Rectangle(xbob3, ybob3, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect3);

                            Rectangle bob1rect4 = new Rectangle(xbob4, ybob4, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect4);

                            Rectangle bob1rect5 = new Rectangle(xbob5, ybob5, 210, 28);
                            e.Graphics.DrawRectangle(blackPen1, bob1rect5);

                            e.Graphics.DrawString(bob1, drawFontArialBold, drawBrush, bob1rect, drawFormatCenter);
                            e.Graphics.DrawString(bob2, drawFontArialBold, drawBrush, bob1rect2, drawFormatCenter);
                            e.Graphics.DrawString(bob3, drawFontArialBold, drawBrush, bob1rect3, drawFormatCenter);
                            e.Graphics.DrawString(bob4, drawFontArialBold, drawBrush, bob1rect4, drawFormatCenter);
                            e.Graphics.DrawString(bob5, drawFontArialBold, drawBrush, bob1rect5, drawFormatCenter);
                            //     y += e.Graphics.MeasureString(bob1, drawFontArialBold).Height;
                            y += bob1rect.Height;
                        }

                        else
                        {

                            y = 98.0F;
                            numberofitems = 0;
                            // NextPage = true;

                            // numberofitemsprintedsofar = 0;
                            e.HasMorePages = true;
                            return;
                        }
                    }
                }




                //if (categoryvalues.Count < 3)
                //{
                //e.Graphics.DrawString(bob1substrings[0], drawFontArial12Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatLeft);
                //    y += e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;

                //e.Graphics.DrawString(bob1substrings[1], drawFontArial12Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatLeft);
                //y += e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;

                //e.Graphics.DrawString(bob1substrings[2], drawFontArial12Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatLeft);
                //y += e.Graphics.MeasureString(bob1, drawFontArial12Regular).Height;
                //}


                //e.Graphics.DrawString(categoryvalues["3"] + ":" + bob3, drawFontArial12Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatLeft);
                //y += e.Graphics.MeasureString(bob1 + bob2, drawFontArial12Regular).Height;


                //if (categoryvalues.Count > 2)
                //{
                //    y += 14;
                //    e.Graphics.DrawString(categoryvalues["3"] + bob1 , drawFontArial12Regular, drawBrush, new RectangleF(x, y, width, height), drawFormatLeft);
                //    y += e.Graphics.MeasureString(categoryvalues["3"] + bob1, drawFontArial12Regular).Height;
                //}
                //}



            }







            // ... and so on

        }

        private int CheckForSizeChange(string value, bool adjustedvalues, int averagesize)
        {
            int newwidth = averagesize;

            if (value == "State:" || value == "Zip:" || value == "City:")
            {
                newwidth = 50;



            }
            else
            {
                newwidth = averagesize;

            }


            return newwidth;

        }

        public class RoutePrinterList
        {
            public string drivername { get; set; }

            public string description { get; set; }
            public string day { get; set; }

            public string businessname { get; set; }
            public string street { get; set; }
            public string notes { get; set; }

            public string city { get; set; }

            public string state { get; set; }

            public string price { get; set; }

            public string zip { get; set; }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            Index = 0;
            numberofitems = 0;
           // bool canConvertCategory = int.TryParse(18, out fontsizeCategory);
            bool canConvert = int.TryParse(txtFontSize.Text, out fontsize);
            if (canConvert != true)
            {
                MessageBox.Show("Font Size value is not valid");
                return;
            }

            categoryvalues = new Dictionary<string, string>();
            reportvalues = new Dictionary<string, string>();
            additionalitems = new SortedDictionary<int, string>();
            cbovalues = new List<string>();

            if (cboBusinessName.Text != "")
            {
                cbovalues.Add(cboBusinessName.Text);

                if (categoryvalues.ContainsKey(cboBusinessName.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;

                }


                categoryvalues.Add(cboBusinessName.Text, "Business Name:");
            }
            if (cboDescription.Text != "")
            {

                if (categoryvalues.ContainsKey(cboDescription.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;

                }


                cbovalues.Add(cboDescription.Text);
                categoryvalues.Add(cboDescription.Text, "Description:");
            }
            if (cboStreet.Text != "")
            {
                if (categoryvalues.ContainsKey(cboStreet.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);

                }


                cbovalues.Add(cboStreet.Text);
                categoryvalues.Add(cboStreet.Text, "Street:");
            }
            if (cboCity.Text != "")
            {
                if (categoryvalues.ContainsKey(cboCity.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }


                cbovalues.Add(cboCity.Text);
                categoryvalues.Add(cboCity.Text, "City:");
            }
            if (cboState.Text != "")
            {
                if (categoryvalues.ContainsKey(cboState.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }



                cbovalues.Add(cboState.Text);
                categoryvalues.Add(cboState.Text, "State:");
            }
            if (cboZip.Text != "")
            {

                if (categoryvalues.ContainsKey(cboZip.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }


                cbovalues.Add(cboZip.Text);
                categoryvalues.Add(cboZip.Text, "Zip:");
            }
            if (cboNotes.Text != "")
            {
                if (categoryvalues.ContainsKey(cboNotes.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }


                cbovalues.Add(cboNotes.Text);
                categoryvalues.Add(cboNotes.Text, "Notes:");
            }
            if (cboPrice.Text != "")
            {
                if (categoryvalues.ContainsKey(cboPrice.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }



                cbovalues.Add(cboPrice.Text);
                categoryvalues.Add(cboPrice.Text, "Price:");
            }

            if (cbovalues.Count == 0)
            {

                MessageBox.Show("There is no order set", "No Order Set Error", MessageBoxButtons.OK);
                return;
            }


            bool isUnique = cbovalues.Distinct().Count() == cbovalues.Count();


            if (!isUnique)
            {
                MessageBox.Show("It looks like some of the order values are the same, these values need to be different in order to be accurate", "Duplicate Order Error", MessageBoxButtons.OK);
                return;
            }

            if (cboDay.Text == "")
            {
                MessageBox.Show("Please choose a day to print from.", "Empty Day Error", MessageBoxButtons.OK);
                return;
            }

            if (cboDrivers.Text == "")
            {
                MessageBox.Show("Please choose a driver.", "Empty Driver Error", MessageBoxButtons.OK);
                return;
            }







            PrintDocument recordDoc = new PrintDocument();


            recordDoc.DefaultPageSettings.Landscape = true;


            float x = 10;
            float y = 5;
            float width = 270.0F; // max width I found through trial and error
            float height = 0F;

            Font drawFontArial12Bold = new Font("Arial", 12, FontStyle.Bold);
            Font drawFontArial10Regular = new Font("Arial", 10, FontStyle.Regular);
            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Set format of string.
            StringFormat drawFormatCenter = new StringFormat();
            drawFormatCenter.Alignment = StringAlignment.Center;
            StringFormat drawFormatLeft = new StringFormat();
            drawFormatLeft.Alignment = StringAlignment.Near;
            StringFormat drawFormatRight = new StringFormat();
            drawFormatRight.Alignment = StringAlignment.Far;
            recordDoc.DocumentName = "Route Report";
            recordDoc.PrintPage += new PrintPageEventHandler(PrintReceiptPage); // function below
            recordDoc.PrintController = new StandardPrintController(); // hides status dialog popup                                            // Comment if debugging 
            PrinterSettings ps = new PrinterSettings();
        //    ps.PrinterName = "Microsoft XPS Document Writer";
            ps.DefaultPageSettings.Landscape = true;
            PrintDialog pdi = new PrintDialog();
            //pdi
            // recordDoc.Print();
            // --------------------------------------
            if (pdi.ShowDialog() == DialogResult.OK)
            {
              //  printPrvDlg.Document = recordDoc;
                //printPrvDlg.Width = 1200;
                //printPrvDlg.Height = 800;
                //pdi.PrinterSettings = ps;
                recordDoc.PrinterSettings = ps;
                //if (pdi.ShowDialog() == DialogResult.OK)
                //{
                //    printPrvDlg.ShowDialog();
                //}
              // string thing =  pdi.Document.PrinterSettings.PrinterName;
                recordDoc.PrinterSettings.PrinterName = pdi.PrinterSettings.PrinterName;
              // pdi.PrinterSettings.PrinterName;
                recordDoc.Print();
            }

                    // Uncomment if debugging - shows dialog instead
           // PrintPreviewDialog printPrvDlg = new PrintPreviewDialog();



            // --------------------------------------

            //   recordDoc.Dispose();
        }

        private void cboDrivers_SelectedValueChanged(object sender, EventArgs e)
        {
            // var updateddays = pRouteList.Where(s => s.drivername == cboDrivers.Text);
            cboDay.Text = "";
            var updateddays = pRouteList.Where(c => c.drivername == cboDrivers.Text).GroupBy(c => c.day).Select(g => g.First());

            cboDay.Items.Clear();
            // cboDay.SelectedValue = "";

            foreach (var item in updateddays)
            {

                cboDay.Items.Add(item.day);
            }
            //   cboDay.SelectedIndex = -1;
            // cboDrivers.SelectedIndex = 0;
        }

        private void chkBusinessName_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkBusinessName.Checked = false;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                counter++;
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }



            if (chkBusinessName.Checked == true)
            {




                cboBusinessName.Enabled = true;
                counter++;

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();



                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());

                }
            }
            else
            {
                cboBusinessName.Enabled = false;
                counter--;


                cboBusinessName.Text = "";
                cboBusinessName.Items.Clear();


                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();


                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
            //var count = MainGroup.Controls.OfType<CheckBox>().Count(x => x.Checked);

        }

        private void chkDescription_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkDescription.Checked = false;
                counter++;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }
            // var count = MainGroup.Controls.OfType<CheckBox>().Count(x => x.Checked);

            //foreach (CheckBox box in MainGroup.Controls.OfType<CheckBox>())
            //{
            //    if (box.Checked)
            //    {
            //        string hi = "hi";
            //    }
            //}




            if (chkDescription.Checked == true)
            {




                cboDescription.Enabled = true;
                counter++;

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();




                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }

            }
            else
            {
                cboDescription.Enabled = false;
                counter--;

                cboDescription.Text = "";
                cboDescription.Items.Clear();


                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
        }

        private void chkStreet_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkStreet.Checked = false;
                counter++;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }


            if (chkStreet.Checked == true)
            {





                cboStreet.Enabled = true;
                counter++;

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();


                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }

            }
            else
            {
                cboStreet.Enabled = false;
                counter--;

                cboStreet.Text = "";
                cboStreet.Items.Clear();

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {


                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }

            }
        }

        private void chkCity_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkCity.Checked = false;
                counter++;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }


            if (chkCity.Checked == true)
            {



                cboCity.Enabled = true;
                counter++;

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
            else
            {
                cboCity.Enabled = false;
                counter--;
                cboCity.Text = "";
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
        }

        private void chkState_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkState.Checked = false;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                counter++;
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }


            if (chkState.Checked == true)
            {




                cboState.Enabled = true;
                counter++;


                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
            else
            {
                cboState.Enabled = false;
                counter--;

                cboState.Text = "";
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
        }

        private void chkZip_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkZip.Checked = false;
                counter++;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }

            if (chkZip.Checked == true)
            {





                cboZip.Enabled = true;
                counter++;


                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
            else
            {
                cboZip.Enabled = false;
                counter--;

                cboZip.Text = "";
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();



                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
        }

        private void chkNotes_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkNotes.Checked = false;
                counter++;
                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();
                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
                return;
            }


            if (chkNotes.Checked == true)
            {




                cboNotes.Enabled = true;
                counter++;


                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
            else
            {
                cboNotes.Enabled = false;
                counter--;

                cboNotes.Text = "";
                cboNotes.Items.Clear();

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
        }

        private void chkPrice_CheckedChanged(object sender, EventArgs e)
        {
            checkboxcount = 0;
            foreach (GroupBox grpbox in MainGroup.Controls.OfType<GroupBox>())
            {
                foreach (CheckBox box in grpbox.Controls.OfType<CheckBox>())
                {
                    if (box.Checked)
                    {

                        checkboxcount++;
                    }
                }
            }

            if (checkboxcount > 5)
            {
                MessageBox.Show("The maximum amount of items for selection is 5 and it looks like there are 5 items already selected.  Please unselect an item if you would like this item in your printed report.", "Too Many Items Error", MessageBoxButtons.OK);
                chkPrice.Checked = false;
                counter++;

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();



                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }


                return;
            }

            if (chkPrice.Checked == true)
            {





                cboPrice.Enabled = true;
                counter++;


                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {

                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
            else
            {
                cboPrice.Enabled = false;
                counter--;

                cboPrice.Text = "";
                cboPrice.Items.Clear();

                cboBusinessName.Items.Clear();
                cboDescription.Items.Clear();
                cboStreet.Items.Clear();
                cboCity.Items.Clear();
                cboState.Items.Clear();
                cboZip.Items.Clear();
                cboNotes.Items.Clear();
                cboPrice.Items.Clear();

                for (i = 1; i <= counter; i++)
                {



                    cboBusinessName.Items.Add(i.ToString());
                    cboDescription.Items.Add(i.ToString());
                    cboStreet.Items.Add(i.ToString());
                    cboCity.Items.Add(i.ToString());
                    cboState.Items.Add(i.ToString());
                    cboZip.Items.Add(i.ToString());
                    cboNotes.Items.Add(i.ToString());
                    cboPrice.Items.Add(i.ToString());
                }
            }
        }

        private void btnPreview_Click(object sender, EventArgs e)

        {

            if (chkBusinessName.Checked && cboBusinessName.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkDescription.Checked && cboDescription.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkStreet.Checked && cboStreet.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkCity.Checked && cboCity.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkState.Checked && cboState.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkZip.Checked && cboZip.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkPrice.Checked && cboPrice.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            if (chkNotes.Checked && cboNotes.Text == "")
            {
                MessageBox.Show("Value checked but no assigned number.  Please fill in the drop down value.", "Missing Value", MessageBoxButtons.OK);
                return;
            }
            Index = 0;
            numberofitems = 0;
           // bool canConvertCategory = int.TryParse(txtCategoryFontSize.Text, out fontsizeCategory);
            bool canConvert = int.TryParse(txtFontSize.Text, out fontsize);
            if (canConvert != true)
            {
                MessageBox.Show("Font Size value is not valid");
                return;
            }

            categoryvalues = new Dictionary<string, string>();
            reportvalues = new Dictionary<string, string>();
            additionalitems = new SortedDictionary<int, string>();
            cbovalues = new List<string>();

            if (cboBusinessName.Text != "")
            {
                cbovalues.Add(cboBusinessName.Text);

                if (categoryvalues.ContainsKey(cboBusinessName.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;

                }


                categoryvalues.Add(cboBusinessName.Text, "Business Name:");
            }
            if (cboDescription.Text != "")
            {

                if (categoryvalues.ContainsKey(cboDescription.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;

                }


                cbovalues.Add(cboDescription.Text);
                categoryvalues.Add(cboDescription.Text, "Description:");
            }
            if (cboStreet.Text != "")
            {
                if (categoryvalues.ContainsKey(cboStreet.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);

                }


                cbovalues.Add(cboStreet.Text);
                categoryvalues.Add(cboStreet.Text, "Street:");
            }
            if (cboCity.Text != "")
            {
                if (categoryvalues.ContainsKey(cboCity.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }


                cbovalues.Add(cboCity.Text);
                categoryvalues.Add(cboCity.Text, "City:");
            }
            if (cboState.Text != "")
            {
                if (categoryvalues.ContainsKey(cboState.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }



                cbovalues.Add(cboState.Text);
                categoryvalues.Add(cboState.Text, "State:");
            }
            if (cboZip.Text != "")
            {

                if (categoryvalues.ContainsKey(cboZip.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }


                cbovalues.Add(cboZip.Text);
                categoryvalues.Add(cboZip.Text, "Zip:");
            }
            if (cboNotes.Text != "")
            {
                if (categoryvalues.ContainsKey(cboNotes.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }


                cbovalues.Add(cboNotes.Text);
                categoryvalues.Add(cboNotes.Text, "Notes:");
            }
            if (cboPrice.Text != "")
            {
                if (categoryvalues.ContainsKey(cboPrice.Text))
                {
                    MessageBox.Show("Uh-Oh, it looks like some of the values for the order are the same.  Lets go ahead and change that so they are all different and try again.", "Same Value Error", MessageBoxButtons.OK);
                    return;
                }



                cbovalues.Add(cboPrice.Text);
                categoryvalues.Add(cboPrice.Text, "Price:");
            }

            if (cbovalues.Count == 0)
            {

                MessageBox.Show("There is no order set", "No Order Set Error", MessageBoxButtons.OK);
                return;
            }


            bool isUnique = cbovalues.Distinct().Count() == cbovalues.Count();


            if (!isUnique)
            {
                MessageBox.Show("It looks like some of the order values are the same, these values need to be different in order to be accurate", "Duplicate Order Error", MessageBoxButtons.OK);
                return;
            }

            if (cboDay.Text == "")
            {
                MessageBox.Show("Please choose a day to print from.", "Empty Day Error", MessageBoxButtons.OK);
                return;
            }

            if (cboDrivers.Text == "")
            {
                MessageBox.Show("Please choose a driver.", "Empty Driver Error", MessageBoxButtons.OK);
                return;
            }







            PrintDocument recordDoc = new PrintDocument();


            recordDoc.DefaultPageSettings.Landscape = true;


            float x = 10;
            float y = 5;
            float width = 270.0F; // max width I found through trial and error
            float height = 0F;

            Font drawFontArial12Bold = new Font("Arial", 12, FontStyle.Bold);
            Font drawFontArial10Regular = new Font("Arial", 10, FontStyle.Regular);
            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Set format of string.
            StringFormat drawFormatCenter = new StringFormat();
            drawFormatCenter.Alignment = StringAlignment.Center;
            StringFormat drawFormatLeft = new StringFormat();
            drawFormatLeft.Alignment = StringAlignment.Near;
            StringFormat drawFormatRight = new StringFormat();
            drawFormatRight.Alignment = StringAlignment.Far;
            recordDoc.DocumentName = "Populating Route";
            recordDoc.PrintPage += new PrintPageEventHandler(PrintReceiptPage); // function below
            recordDoc.PrintController = new StandardPrintController(); // hides status dialog popup                                            // Comment if debugging 
            PrinterSettings ps = new PrinterSettings();
            ps.PrinterName = "Microsoft XPS Document Writer";
            ps.DefaultPageSettings.Landscape = true;
          //  PrintDialog pdi = new PrintDialog();
            //pdi
            // recordDoc.Print();
          //  // --------------------------------------
           // pdi.ShowDialog();
            // Uncomment if debugging - shows dialog instead
            PrintPreviewDialog printPrvDlg = new PrintPreviewDialog();
            printPrvDlg.Document = recordDoc;
            //printPrvDlg.Width = 1200;
            //printPrvDlg.Height = 800;
           // pdi.PrinterSettings = ps;
            recordDoc.PrinterSettings = ps;

            printPrvDlg.ShowDialog();
         //   if (pdi.ShowDialog() == DialogResult.OK)
          //  {
         //       printPrvDlg.ShowDialog();
          //  }
            //  recordDoc.PrinterSettings.PrinterName = pdi.

            //        recordDoc.Print();


            // --------------------------------------

            //   recordDoc.Dispose();
        }
    }  
    }

