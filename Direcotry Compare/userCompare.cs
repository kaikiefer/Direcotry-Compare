using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Direcotry_Compare
{
    public partial class userCompare : Form
    {

        //Create variables
        string[] user1 = new string[500];
        string[] user2 = new string[500];
        string[] similar = new string[500];
        string[] dissimilar = new string[500];
        string userName1 = "";
        string userName2 = "";
        string strFile = null;
        public userCompare()
        {
            InitializeComponent();

            //Set up screen for first use
            this.Width = 337;
            this.Height = 232;
            label2.Visible = true;
            label3.Visible = true;
            textBox1.Visible = true;
            textBox2.Visible = true;
            groupBox1.Visible = true;
            button1.Visible = true;
            label1.Visible = false;
            listBox1.Visible = false;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //Set to loading cursor
            Cursor.Current = Cursors.WaitCursor;

            //Change view
            label2.Visible = false;
            label3.Visible = false;
            textBox1.Visible = false;
            textBox2.Visible = false;
            groupBox1.Visible = false;
            button1.Visible = false;
            label1.Visible = true;
            listBox1.Visible = true;
            this.Width = 1371;
            this.Height = 711;

            //Instantiate AD class
            ADActions ad = new ADActions();

            //Get data from user inputs
            userName1 = textBox1.Text;
            userName2 = textBox2.Text;

            //Renable the boxes
            label1.Text = userName1;
            label4.Text = userName2;

            //Poll for user 1 AD groups
            try
            {
                user1 = ad.getUserGroups(userName1);
            }
            catch
            {
                //unable to poll for user groups
            }

            //Add groups to user 1 box
            for (int i = 0; i < user1.Length; i++)
            {
                if (user1[i] != null)
                {
                    listBox1.Items.Add(user1[i]);
                }
                else
                {
                    //lol nothing
                }
            }

            //Poll for user 2 AD groups
            try
            {
                user2 = ad.getUserGroups(userName2);
            }
            catch
            {
                //unable to poll for user groups
            }

            //Add groups to user 2 box
            for (int i = 0; i < user2.Length; i++)
            {
                if (user2[i] != null)
                {
                    listBox2.Items.Add(user2[i]);
                }
                else
                {
                    //lol nothing
                }
            }

            //Set up counters for similar and dissimilar arrays
            int s = 0;
            int d = 0;

            //Find similar groups
            for (int i = 0; i < user1.Length; i++)
            {
                for (int j = 0; j < user2.Length; j++)
                {
                    if (user1[i] == user2[j])
                    {
                        similar[s] = user1[i];
                        s++;
                    }
                }
            }

            //Find dissimilar groups from user 1
            for (int i = 0; i < user1.Length; i++)
            {
                if (similar.Contains(user1[i]))
                {
                    //skip it
                }
                else
                {
                    dissimilar[d] = user1[i];
                    d++;
                }
            }

            //find dissimilar groups from user 2
            for (int i = 0; i < user2.Length; i++)
            {
                if (similar.Contains(user2[i]) || dissimilar.Contains(user2[i]))
                {
                    //skip it
                }
                else
                {
                    dissimilar[d] = user2[i];
                    d++;
                }
            }

            //Add similar groups to list box
            for (int i = 0; i < similar.Length; i++)
            {
                if (similar[i] != null)
                {
                    listBox3.Items.Add(similar[i]);
                }
                else
                {
                    //lol nothing
                }
            }

            //Add dissimilar groups to listbox
            for (int i = 0; i < dissimilar.Length; i++)
            {
                if (dissimilar[i] != null)
                {
                    listBox4.Items.Add(dissimilar[i]);
                }
                else
                {
                    //lol nothing
                }
            }

            //Count groups in each listbox
            int user1Count = listBox1.Items.Count;
            int user2Count = listBox2.Items.Count;
            int similarCount = listBox3.Items.Count;
            int dissimilarCount = listBox4.Items.Count;

            //Show group counts on the right
            label9.Text = user1Count.ToString();
            label10.Text = user2Count.ToString();
            label13.Text = similarCount.ToString();
            label15.Text = dissimilarCount.ToString();

            //Change user names on the right
            label8.Text = userName1;
            label11.Text = userName2;

            //Set to default cursor
            Cursor.Current = Cursors.Default;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //Reset
            textBox1.Text = "";
            textBox2.Text = "";
            this.Width = 337;
            this.Height = 232;
            label2.Visible = true;
            label3.Visible = true;
            textBox1.Visible = true;
            textBox2.Visible = true;
            groupBox1.Visible = true;
            button1.Visible = true;
            label1.Visible = false;
            listBox1.Visible = false;
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            label17.Visible = false;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //Set to loading cursor
            Cursor.Current = Cursors.WaitCursor;

            //Export to excel
            label17.Visible = true;
            label17.Text = "Exporting...";

            try
            {

                Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook oWorkBook;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;

                oExcel.Visible = false;
                oWorkBook = (Microsoft.Office.Interop.Excel._Workbook)(oExcel.Workbooks.Add(Missing.Value));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWorkBook.ActiveSheet;

                oSheet.Cells[1, 1] = "User: " + userName1;
                oSheet.Cells[1, 2] = "User: " + userName2;
                oSheet.Cells[1, 3] = "Similar Groups";
                oSheet.Cells[1, 4] = "Dissimilar Groups";

                //User 1
                for (int j = 0; j < user1.Length; j++)//inserting AD groups into column one
                {
                    oSheet.Cells[j + 2, 1] = user1[j];
                }

                //User 2
                for (int j = 0; j < user2.Length; j++)//inserting AD groups into column one
                {
                    oSheet.Cells[j + 2, 2] = user2[j];
                }

                //Similar
                for (int j = 0; j < similar.Length; j++)//inserting AD groups into column one
                {
                    oSheet.Cells[j + 2, 3] = similar[j];
                }

                //Dissimilar
                for (int j = 0; j < dissimilar.Length; j++)//inserting AD groups into column one
                {
                    oSheet.Cells[j + 2, 4] = dissimilar[j];
                }

                strFile = "Tranquility AD Compare Export for " + userName1 + " and " + userName2 + " " + System.DateTime.Now.Ticks.ToString() + ".xls";
                oWorkBook.SaveAs("c:\\Users\\Public\\Documents\\" + strFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, null, null, null);
                oWorkBook.Close(null, null, null);
                oExcel.Workbooks.Close();
                oExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWorkBook);
                oSheet = null;
                oWorkBook = null;
                oExcel = null;
                System.Diagnostics.Process.Start("c:\\Users\\Public\\Documents\\" + strFile);

                //Notify of completion
                label17.Text = "Complete";
                label17.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                label17.Text = "ERROR";
                label10.ForeColor = Color.Red;
            }

            //Set to default cursor
            Cursor.Current = Cursors.Default;
        }

        private void listBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {

            //Get currently selected similar
            var currentItem = listBox3.SelectedItem;

            //Select the item on the users
            listBox1.SelectedItem = currentItem;
            listBox2.SelectedItem = currentItem;

            //Clear the items on the dissimilar list
            listBox4.SelectedItem = null;
        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Get current dissimilar group
            var currentItem = listBox4.SelectedItem;

            //Clear similar
            listBox3.SelectedItem = null;

            //select if in the user 1 section
            try
            {
                listBox1.SelectedItem = currentItem;
            }
            catch
            {
                listBox1.SelectedItem = null;
            }

            //Select if in the user 2 section
            try
            {
                listBox2.SelectedItem = currentItem;
            }
            catch
            {
                listBox2.SelectedItem = null;
            }
        }
    }
}
