using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace TxtFileConvert
{
    public partial class Historical_Data : Form
    {

        //create string array variable to capture selections
        //string[] fileArray = new string[];
        public static DataTable fileArray = new DataTable();

        public Historical_Data()
        {
            InitializeComponent();

            DataColumn InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "FileName";

            fileArray.Columns.Add(InputColumn);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Stream myStream = null;
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }


            }

            //DataTable outputtable = Form1.OutputDataTable1();
            //loop through selection populating global string array varaible
            DataRow fileArrayRow;

            foreach (String file in openFileDialog1.FileNames)
            {

                fileArrayRow = fileArray.NewRow();
                fileArrayRow[0] = file;
                fileArray.Rows.Add(fileArrayRow);
                //MessageBox.Show(fileArray.Rows[0][0].ToString());
            }

            //stop here

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //loop through each file in the string array, but continually add to the same data table

            

            //Establish integer to determine creation of new rows
            //one the beginning of each loop, set integer to 0
            //if 0, input new item string
            //trim item string to take out characters, convert to integer
            //else, determine if next item string, after trim, is less than or equal to current
            //if less than or equal, create new row
            //else add to same row
        }
    }
}
