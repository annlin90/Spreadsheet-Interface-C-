using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CptS321;
using SpreadsheetEngine;
using System.Xml;
using System.IO;

namespace Spreadsheet_LinA
{
    public partial class Form1 : Form
    {   
        Spreadsheet A = new Spreadsheet(50, 26);
        
        public Form1()
        {
            InitializeComponent();
            A.CellPropertyChanged += OnCellPropertyChanged;      
        }
         
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            dataGridView1.Columns.Add("1", "A");
            dataGridView1.Columns.Add("2", "B");
            dataGridView1.Columns.Add("3", "C");
            dataGridView1.Columns.Add("4", "D");
            dataGridView1.Columns.Add("5", "E");
            dataGridView1.Columns.Add("6", "F");
            dataGridView1.Columns.Add("7", "G");
            dataGridView1.Columns.Add("8", "H");
            dataGridView1.Columns.Add("9", "I");
            dataGridView1.Columns.Add("10", "J");
            dataGridView1.Columns.Add("11", "K");
            dataGridView1.Columns.Add("12", "L");
            dataGridView1.Columns.Add("13", "M");
            dataGridView1.Columns.Add("14", "N");
            dataGridView1.Columns.Add("15", "O");
            dataGridView1.Columns.Add("16", "P");
            dataGridView1.Columns.Add("17", "Q");
            dataGridView1.Columns.Add("18", "R");
            dataGridView1.Columns.Add("19", "S");
            dataGridView1.Columns.Add("20", "T");
            dataGridView1.Columns.Add("21", "U");
            dataGridView1.Columns.Add("22", "V");
            dataGridView1.Columns.Add("23", "W");
            dataGridView1.Columns.Add("24", "X");
            dataGridView1.Columns.Add("25", "Y");
            dataGridView1.Columns.Add("26", "Z");

            for (int i = 0; i < 50; i++)
            {
                dataGridView1.Rows.Add(" ");
            }
            foreach (DataGridViewRow dGVRow in this.dataGridView1.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            
            //This resizes the width of the row headers to fit the numbers
            this.dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }
        private void OnCellPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            dataGridView1.Rows[((Cell)sender).Rows].Cells[((Cell)sender).Columns].Value = ((Cell)sender).Values;
        }
        private void button1_Click(object sender, EventArgs e)
        {

            Random rand = new Random();
            for (int i = 0; i < 50; i++)
            {
                int a = rand.Next(0, 49);  //sets 50 random cells to Elementary! this is possible # of rows it can appear on
                int b = rand.Next(2, 25);  //this is possible # of columns it can appear on 

                A.Arrays[a, b].Texter = "Elementary!"; //sets the text to random rows and column in Spreadsheet array
              
                //then set those values onto dataGridView
            }
            
            for (int i = 0; i < 50; i++) {
                string num = (i + 1).ToString();
                A.Arrays[i, 1].Texter= "This is cell B" + num; //same thing, except set B column to this text
               
            }
             
            for (int i = 0; i < 50; i++) {
                string num = (i + 1).ToString();
                A.Arrays[i, 0].Texter = "=B" + num; //same thing, except set A column to the same text, making it so 
                //that A column contents will be the same as B columns contents
               
            }
        }

        //code for save button
        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog SFD = new SaveFileDialog();
            SFD.DefaultExt = "xml";
            if (SFD.ShowDialog() == DialogResult.OK)
            {
                FileStream fsStream = new FileStream(SFD.FileName, FileMode.Create, FileAccess.Write);
                A.XMLSave(fsStream);
                fsStream.Dispose();
            }
            else
            {
                MessageBox.Show("Error in path");
            }
        }

        //code for load button
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog One = new OpenFileDialog();
            XmlDocument First = new XmlDocument();
            if (One.ShowDialog() == DialogResult.OK)
            {
                FileStream myStream = new FileStream(One.FileName, FileMode.Open, FileAccess.Read);
                A.XMLLoad(myStream);
                myStream.Dispose();
                First.Load(One.FileName);
            }
        }
    }
 }

