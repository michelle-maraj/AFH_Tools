using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Net.Mime;
using System.Text.RegularExpressions;

namespace AFH_Tools
{
    
    public partial class AFH : Form
    {
        private const string FileName = "Notepad.exe";

        //string sharepoint_milli = @"https://4thtest.sharepoint.com/Operations/St%20Jude/Forms/AllItems.aspx?id=%2FOperations%2FSt%20Jude%2FMilli%20SOPs";
        //string sharepoint_hp = @"https://4thtest.sharepoint.com/:f:/g/Operations/Er0My5UqB75Gnu6pgG_LgFkBgqL8AAV_UBK9z-UPAnWPbg?e=6PD6UV";
        string openair = @"https://www.openair.com/index.pl";
        string operations = @"https://4thtest.sharepoint.com/Operations/SitePages/Home.aspx";
        string servicenow = @"https://4thsource.service-now.com/nav_to.do?uri=%2Fhome.do";
        string five9 = @"https://login.five9.com/";
        string googlevoice = @"https://voice.google.com/messages";
        string fxwell = @"https://fxwell.com/Account/Login?ReturnUrl=%2FPortal%2FHome%2FIndex";
        string esiepa = @"https://accessps.express-scripts.com/epa/epa.html";
        string yammer = @"https://www.yammer.com/4thsource.com/#/threads/company?type=general";
        string staffhub = @"https://staffhub.office.com/app";

        string link = "";
        //production links
        string milli = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\milli\";
        string helpdesk = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\helpdesk\";

        string notes_file = @"C:\Users\MichelleMaraj\Desktop\st jude sop\notes\";
        int counter = 0;
        string getfile ="";

        public AFH()
        {
            InitializeComponent();
        }
        private static void linkMethod(string link)
        {
            try
            {
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception ex)
            {

                MessageBox.Show("error" + ex);

            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = servicenow;
            linkMethod(link);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = openair;
            linkMethod(link);
        }

        private void linkLabel_sharepoint_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = operations;
            linkMethod(link);
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = five9;
            linkMethod(link);
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = googlevoice;
            linkMethod(link);
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = fxwell;
            linkMethod(link);
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = esiepa;
            linkMethod(link);
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = yammer;
            linkMethod(link);
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = staffhub;
            linkMethod(link);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Form2 secondForm;
                secondForm = new Form2();
                secondForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("error" + ex);
            }
        }

        private string directories(string dir)
        {
            if (checkBox1.Checked)
            {
                dir = milli;
            }
            else if (checkBox2.Checked)
            {
                dir = helpdesk;
            }
           

            return dir;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                listBox1.Items.Clear();
                richTextBox1.Clear();
                string[] files = Directory.GetFiles(milli);
                foreach (string file in files)
                {
                    listBox1.Items.Add(Path.GetFileName(file));

                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                try
                {
                    listBox1.Items.Clear();
                    richTextBox1.Clear();
                    string[] files = Directory.GetFiles(helpdesk);
                    foreach (string file in files)
                    {
                        listBox1.Items.Add(Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
        }

        private void opensj_Click(object sender, EventArgs e)
        {
            try
            {
                String dir = "";
                axAcroPDF1.Visible = true;
                webBrowser1.Visible = false;
                dir = directories(dir);

                string curItem = listBox1.SelectedItem.ToString();
                MessageBox.Show(curItem);
                //now to open file
                string new_file = dir + curItem;

                if (curItem.ToString().ToLower().Contains(".pdf"))
                {
                    axAcroPDF1.src = new_file;
                    //Html_file.Text = file;
                    //Uri uri = new Uri(new_file);
                    //webBrowser1.Navigate(uri);
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("adobe reader not installed" + ex);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                listBox1.SelectedItems.Clear();
                for (int i = listBox1.Items.Count - 1; i >= 0; i--)
                {
                    if (listBox1.Items[i].ToString().ToLower().Contains(textBox1.Text.ToLower()))
                    {
                        listBox1.SetSelected(i, true);

                    }
                    else if (!listBox1.Items[i].ToString().ToLower().Contains(textBox1.Text.ToLower()))
                    {
                        listBox1.Items.RemoveAt(i);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("cannot find file" + ex);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document WordDoc = new Microsoft.Office.Interop.Word.Document();
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            WordDoc.Close();
            
            axAcroPDF1.Visible = false;

            ap.Quit();
            webBrowser1.Visible = false;
            if (checkBox1.Checked)
            {
                try
                {
                    listBox1.Items.Clear();
                    richTextBox1.Clear();
                    string[] files = Directory.GetFiles(milli);
                    foreach (string file in files)
                    {
                        listBox1.Items.Add(Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
            else if (checkBox2.Checked)
            {
                try
                {
                    listBox1.Items.Clear();
                    richTextBox1.Clear();
                    string[] files = Directory.GetFiles(helpdesk);
                    foreach (string file in files)
                    {
                        listBox1.Items.Add(Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Form3 thirdForm;
                thirdForm = new Form3();
                thirdForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("error" + ex);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(FileName);

        }
        string NextFile(string path, ref int counter)
        {
            var filePath= "";
            if ((filePath = Directory.EnumerateFiles(path).Skip(counter).FirstOrDefault()) != null)
            {
                counter++;
            }
            else
            {
                counter = 0;
                filePath = Directory.EnumerateFiles(path).Skip(counter).FirstOrDefault();
                counter++;
            }
            return filePath;
        }

       

        private void button_forward_Click(object sender, EventArgs e)
        {

            getfile = NextFile(notes_file, ref counter);
            //MessageBox.Show(getfile);
            FileStream inFile = new FileStream(getfile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamReader reader = new StreamReader(inFile);
            string text = reader.ReadToEnd();
            richTextBox1.Text = text;
            reader.Close();
            inFile.Close();

        }

        private void button_back_Click(object sender, EventArgs e)
        {
            string getcurrentfile = Path.GetFileName(getfile);
            System.Diagnostics.Process.Start(notes_file + getcurrentfile);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string getcurrentfile = Path.GetFileName(getfile);
            MessageBox.Show(getcurrentfile + "will be deleted");
            File.Delete(notes_file + getcurrentfile);
        }
    }
}
