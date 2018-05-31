using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Net.Mime;
using System.Text.RegularExpressions;

namespace AFH_Tools
{
    public partial class Form3 : Form
    {
        //store production docx files non production
        string milli_sop = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\Milli SOPs\";
        string hp_sop = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\HelpDesk SOPs\";

        //production links
        string helpdesk = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\helpdesk\";
        string milli = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\milli\";

        public Form3()
        {
            InitializeComponent();
            
        }

        private string directories(string dir)
        {
            if (checkBox1.Checked)
                {
                    dir = hp_sop;
                }
                else if (checkBox2.Checked)
                {
                    dir = milli_sop;
                }


                return dir;
            
        }
        private string newdirectories(string dir)
        {
            
                if (checkBox1.Checked)
                {
                    dir = helpdesk;
                }
                else if (checkBox2.Checked)
                {
                    dir = milli;
                }


                return dir;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                String dir_docx = "";
                dir_docx = directories(dir_docx);

                openFileDialog1.InitialDirectory = dir_docx;

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox1.Text = openFileDialog1.FileName;
                }

                openFileDialog1.Dispose();
            }
            catch (Exception ex)
            {

                MessageBox.Show("error" + ex);
            }

        }
     

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document WordDoc = new Microsoft.Office.Interop.Word.Document();

                string file = textBox1.Text;
                ap.Visible = true;
                WordDoc = ap.Documents.Open("\"" + file + "\"");
                //WordDoc.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show("error" + ex);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Document WordDoc = new Microsoft.Office.Interop.Word.Document();
                Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();

                string file = textBox1.Text;
                string dir = "";
                var fileName = openFileDialog1.FileName;
                WordDoc = ap.Documents.Open(fileName);

                ExportMethod(WordDoc, dir, fileName);
                WordDoc.Close();
                ap.Quit();

            }
            catch (Exception ex)
            {

                MessageBox.Show("error" + ex);
            }

            //System.IO.File.Copy(fileName, Path.Combine(Path.GetDirectoryName(fileName), newdirectories(dir)+  Path.GetFileNameWithoutExtension(fileName) + ".pdf"), true);
            //destFile = System.IO.Path.Combine(newdirectories(dir), fileName);
            //sourceFile = file;
            //System.IO.File.Move(sourceFile, destFile);
        }

        private void ExportMethod(Document WordDoc, string dir, string fileName)
        {
            try
            {
                WordDoc.ExportAsFixedFormat(newdirectories(dir) + Path.GetFileNameWithoutExtension(fileName) + ".pdf", ExportFormat: WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {

                MessageBox.Show("error" + ex);
            }
        }
    }
}
