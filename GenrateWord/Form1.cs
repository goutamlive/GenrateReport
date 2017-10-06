using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace GenrateWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> replaceDict = new Dictionary<string, string>();
            replaceDict.Add("Name", textBox1.Text.Trim());
            replaceDict.Add("Address", textBox2.Text);
            replaceDict.Add("PhoneNumber", textBox3.Text);
            replaceDict.Add("City", textBox4.Text);
            replaceDict.Add("Country", textBox5.Text.Trim());
            replaceDict.Add("TempletPath", textBox2.Text.Trim());
            label7.Text = "WordConvert Started - Time " + DateTime.Now.ToLongTimeString();
            WordConvert(replaceDict);
            label7.Text = "WordConvert Finished - Time " + DateTime.Now.ToLongTimeString();
        }

        static void WordConvert(Dictionary<string, string> replaceDict)
        {
            Console.WriteLine("WordConvert Started");
            Object oMissing = System.Reflection.Missing.Value;
            String defaultPath;
            String userName = Regex.Replace(System.Security.Principal.WindowsIdentity.GetCurrent().Name, ".*\\\\(.*)", "$1", RegexOptions.None);
            bool hasValue = replaceDict.TryGetValue("templetPath", out defaultPath);
            Object oTemplatePath;
            if (hasValue)
            {
                oTemplatePath = defaultPath;
            }
            else
            {
                oTemplatePath = "C:\\Users\\"+ userName + "\\Desktop\\Intro.dotx";
            }

            Console.WriteLine(oTemplatePath);

            word.Application wordApp = new word.Application();
            word.Document wordDoc = new word.Document();

            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (word.Field myMergeField in wordDoc.Fields)
            {

                word.Range rngFieldCode = myMergeField.Code;

                String fieldText = rngFieldCode.Text;

                Console.WriteLine(fieldText);

                // ONLY GETTING THE MAILMERGE FIELDS

                if (fieldText.StartsWith(" MERGEFIELD"))
                {

                    // THE TEXT COMES IN THE FORMAT OF

                    // MERGEFIELD  MyFieldName  \\* MERGEFORMAT

                    // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                    Int32 endMerge = fieldText.IndexOf("\\");

                    Int32 fieldNameLength = fieldText.Length - endMerge;

                    String fieldName = fieldText.Substring(11, endMerge - 11);

                    // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE

                    fieldName = fieldName.Trim();
                    // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//

                    // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE

                    if (fieldName == "Name")
                    {
                        Console.WriteLine("Found Name");
                        myMergeField.Select();

                        wordApp.Selection.TypeText(replaceDict["Name"]);

                    }

                    if (fieldName == "Address")
                    {
                        Console.WriteLine("Found Address");
                        myMergeField.Select();

                        wordApp.Selection.TypeText(replaceDict["Address"]);

                    }

                    if (fieldName == "PhoneNumber")
                    {
                        Console.WriteLine("Found PhoneNumber");
                        myMergeField.Select();

                        wordApp.Selection.TypeText(replaceDict["PhoneNumber"]);

                    }

                    if (fieldName == "City")
                    {
                        Console.WriteLine("Found City");
                        myMergeField.Select();

                        wordApp.Selection.TypeText(replaceDict["City"]);

                    }

                    if (fieldName == "Country")
                    {
                        Console.WriteLine("Found Country");
                        myMergeField.Select();

                        wordApp.Selection.TypeText(replaceDict["Country"]);

                    }

                }

            }
            wordDoc.SaveAs("C:\\Users\\"+ "goukumar" + "\\Desktop\\myfile.doc");
            Console.WriteLine("Converted .... quitting");
            wordApp.Application.Quit();
        }
    }
}
