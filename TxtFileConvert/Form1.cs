using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Collections.Generic;

namespace TxtFileConvert
{
    public partial class Form1 : Form
    {

        public string InputFileName = "";
        public string InputFileNameFullPath = "";
        public DataTable Output = new DataTable();
        public DataTable outputtable1 = new DataTable();
        public string fiscalmonth = "";
        public string fiscalyear = "";
        public string BUKRS = "";
        public string lastDayMonth = "";
        public string CC = "";

        public static DataTable fileArray = new DataTable();



        public Form1()
        {

            InitializeComponent();
            DataColumn InputColumnFile = new DataColumn();
            InputColumnFile.DataType = System.Type.GetType("System.String");
            InputColumnFile.ColumnName = "FileName";

            fileArray.Columns.Add(InputColumnFile);
            //  Historical_Data hisdata = new Historical_Data();
            //hisdata.Activate();
            //hisdata.Show();
            //Application.Run();


            // Bring Create data table for conversion


            //Bring in headers for Output
            Assembly _assembly;
            StreamReader GetCSVColumnHeaders;


            _assembly = Assembly.GetExecutingAssembly();
            GetCSVColumnHeaders = new StreamReader(_assembly.GetManifestResourceStream("TxtFileConvert.TextFile1.txt"));


            char[] delimiter = new char[] { '\t' };

            string[] outputheader = new string[170];
            //string[] columnheaders = GetCSVColumnHeaders.ReadLine().Split(delimiter);

            DataTable InputFile = new DataTable("InputFile");


            //create columns on table InputFile
            DataColumn InputColumn;
            DataColumn OutputColumn = new DataColumn();
            DataRow InputRow;


            DataRow ToCSVRef;
            ToCSVRef = InputFile.NewRow();
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "CSVFileLocation";
            InputFile.Columns.Add(InputColumn);
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "CSVFileColumnName";
            InputFile.Columns.Add(InputColumn);
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "TXTColumnName";
            InputFile.Columns.Add(InputColumn);
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "TXTColumnLocation";
            InputFile.Columns.Add(InputColumn);

            int integer1 = 0;
            while (!GetCSVColumnHeaders.EndOfStream)
            {

                string[] s = GetCSVColumnHeaders.ReadLine().Split(delimiter);


                if (s[3].ToString().Trim() != "")
                {

                    InputRow = InputFile.NewRow();
                    InputRow["CSVFileLocation"] = s[0].ToString();
                    InputRow["CSVFileColumnName"] = s[1].ToString();
                    InputRow["TXTColumnName"] = s[2].ToString();
                    InputRow["TXTColumnLocation"] = s[3].ToString();
                    InputFile.Rows.Add(InputRow);

                }
                //outputheader[integer1] = s[1];
                //MessageBox.Show(s[1].ToString());
                integer1++;
                //ToCSVRef[s[1].ToString()] = s[2].ToString();


                OutputColumn = new DataColumn(s[1].ToString(), System.Type.GetType("System.String"));
                outputtable1.Columns.Add(OutputColumn);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable datatable1 = new DataTable();
            DataTable outputtable1 = new DataTable();
            //Bring in TXT Sales File
            StreamReader streamreader = new StreamReader(InputFileName);


            //Bring Create data table for conversion
            DataTable Output = new DataTable();

            //Bring in headers for Output
            Assembly _assembly;
            StreamReader GetCSVColumnHeaders;


            _assembly = Assembly.GetExecutingAssembly();
            GetCSVColumnHeaders = new StreamReader(_assembly.GetManifestResourceStream("TxtFileConvert.TextFile2.txt"));


            char[] delimiter = new char[] { '\t' };

            string[] outputheader = new string[170];
            //string[] columnheaders = GetCSVColumnHeaders.ReadLine().Split(delimiter);

            DataTable InputFile = new DataTable("InputFile");


            //create columns on table InputFile
            DataColumn InputColumn;
            DataColumn OutputColumn = new DataColumn();
            DataRow InputRow;


            DataRow ToCSVRef;
            ToCSVRef = InputFile.NewRow();
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "CSVFileLocation";
            InputFile.Columns.Add(InputColumn);
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "CSVFileColumnName";
            InputFile.Columns.Add(InputColumn);
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "TXTColumnName";
            InputFile.Columns.Add(InputColumn);
            InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "TXTColumnLocation";
            InputFile.Columns.Add(InputColumn);

            int integer1 = 0;
            while (!GetCSVColumnHeaders.EndOfStream)
            {

                string[] s = GetCSVColumnHeaders.ReadLine().Split(delimiter);


                if (s[3].ToString().Trim() != "")
                {

                    InputRow = InputFile.NewRow();
                    InputRow["CSVFileLocation"] = s[0].ToString();
                    InputRow["CSVFileColumnName"] = s[1].ToString();
                    InputRow["TXTColumnName"] = s[2].ToString();
                    InputRow["TXTColumnLocation"] = s[3].ToString();
                    InputFile.Rows.Add(InputRow);

                }
                //outputheader[integer1] = s[1];
                //MessageBox.Show(s[1].ToString());
                integer1++;
                //ToCSVRef[s[1].ToString()] = s[2].ToString();


                OutputColumn = new DataColumn(s[1].ToString(), System.Type.GetType("System.String"));
                outputtable1.Columns.Add(OutputColumn);
            }

            //take contents of streamreader and output to outputtable1
            DataRow OutPutRow;
            int i = 0;

            //loop through txt file
            while (!streamreader.EndOfStream)
            {
                //split streamreader line into string array using /t delimiter
                string[] inputstring = streamreader.ReadLine().Split(delimiter);

                //determine if a new row is needed or this is a row with a contribution margin
                if (inputstring[6].ToString() == "CM10010ITM")
                {
                    //if it is CM10010ITM this represents a new row, create a new row on outputtable
                    OutPutRow = outputtable1.NewRow();
                    //go through each element on streamreader and put it into the corressponding column in the output table
                    foreach (DataRow rowloop in InputFile.Rows)
                    {
                        if (rowloop[3].ToString() == "7")
                        { }
                        else
                        {
                            try
                            { OutPutRow[Convert.ToInt32(rowloop[0].ToString())] = inputstring[Convert.ToInt32(rowloop[3].ToString())].ToString(); }

                            catch (Exception eE)
                            {
                                MessageBox.Show(rowloop[0].ToString() + rowloop[3].ToString());
                            }


                        }

                    }

                    outputtable1.Rows.Add(OutPutRow);
                    i++;
                }
            }


            //dataGridView1.DataSource = outputtable1;
            outputtable1.ExportToExcel("excel.xlsx");



        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Stream myStream = null;
            openFileDialog1.InitialDirectory = "c:\\";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
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
                MessageBox.Show(fileArray.Rows[0][0].ToString());
            }



        }
        //build out the output data table
        private void button3_Click(object sender, EventArgs e)
        {

            //Bring in TXT Sales File
            StreamReader streamreaderTXT = new StreamReader(InputFileNameFullPath);
            DataRow ExtractRow;

            //identify "tab" as the column delimiter in the txt file
            char[] delimiter = new char[] { '\t' };

            using (DataTable dtx = OutputDataTable1())
            {


                if (InputFileName == "")
                    MessageBox.Show("Please Select File to Convert");

                else
                {
                    ExtractRow = dtx.NewRow();

                    while (!streamreaderTXT.EndOfStream)
                    {
                        //split streamreader line into string array using /t delimiter
                        string[] inputstring = streamreaderTXT.ReadLine().Split(delimiter);

                        // int rowcounter = 0;
                        //string column6 = "";
                        int CMINT;

                        //get contribution margin identifier for conversion to number


                        //determine if a new row is needed or this is a row with a contribution margin
                        //add to take into account "CFCH00620" for transactional files
                        if ((inputstring[6].ToString() == "Position") || (inputstring[6].ToString() == "CFCH00620"))
                            continue;

                        //MessageBox.Show(inputstring[6].ToString());
                        //CMINT = Convert.ToInt32(inputstring[6].ToString().Substring(2, 5));
                        //MessageBox.Show(CMINT.ToString());

                        //need to change this.  It should start with the next row.  the CM at column 6 should become a double (take off the "CM" and the "ITM/SUM") - Then if the next row is =<, start a new row
                        //this else forces the creation of a new row - need to figure out how to identify that
                        else if (inputstring[6].ToString() == "CM10010ITM")
                        {
                            //creation of new row
                            ExtractRow = dtx.NewRow();

                            //initiate each field to "0" - Qlik requirement
                            for (int i = 0; i < ExtractRow.Table.Columns.Count; i++)
                            {

                                ExtractRow[i] = "0";

                            }



                            //take # in column 41, and post it in the corresponding contribution margin column in the new row
                            ExtractRow[inputstring[6].ToString()] = inputstring[41].ToString();

                            ExtractRow["BUDAT - Posting Date"] = lastDayMonth;
                            ExtractRow["FDAT - Invoice Date"] = lastDayMonth;
                            ExtractRow["BURKS - Company Code"] = CC;  //need to figure this piece out?
                            ExtractRow["VERSI - Version"] = inputstring[0].ToString();
                            ExtractRow["CPLYEAR - Planning Year"] = inputstring[1].ToString();
                            ExtractRow["0FISCYEAR-Planning Year"] = fiscalyear;
                            ExtractRow["WWPER - Fiscal Period"] = fiscalmonth;
                            ExtractRow["CCOMPANY - Company"] = inputstring[4].ToString();
                            ExtractRow["VBUND - Partner Company"] = inputstring[5].ToString();
                            ExtractRow["CFCH002620 - Item"] = "";
                            ExtractRow["KOKRS - Controlling Area"] = inputstring[7].ToString();
                            ExtractRow["PRCTR - Profit Center"] = inputstring[8].ToString();
                            ExtractRow["CPPRCTR - Partner Profit Center"] = inputstring[9].ToString();
                            ExtractRow["WWBRN - Branch/Industry"] = inputstring[10].ToString();
                            ExtractRow["WWPRG - Product Group"] = inputstring[11].ToString();
                            ExtractRow["WWPPG - Partner Product Group"] = inputstring[12].ToString();
                            ExtractRow["WWART - Material Number"] = inputstring[13].ToString();
                            ExtractRow["KUNWE - Ship-To (local)"] = inputstring[14].ToString();
                            ExtractRow["KNDNR - Sold-To (local)"] = inputstring[15].ToString();
                            ExtractRow["KUNRE - Bill-To (local)"] = inputstring[16].ToString();
                            ExtractRow["KUNRG - Payer (local)"] = inputstring[17].ToString();
                            ExtractRow["WWKUN - Ship-To Final (local)"] = inputstring[18].ToString();
                            ExtractRow["WWLWE - Country (Ship-To)"] = inputstring[19].ToString();
                            ExtractRow["LAND1 - Country (Sold-To)"] = inputstring[20].ToString();
                            ExtractRow["WWLRE - Country (Bill-To)"] = inputstring[21].ToString();
                            ExtractRow["WWLRG - Country (Payer)"] = inputstring[22].ToString();
                            ExtractRow["WWFCU - Country (Ship-To Final)"] = inputstring[23].ToString();
                            ExtractRow["KSTRG - Cost object"] = inputstring[24].ToString();
                            ExtractRow["WWKAF - Sales order"] = inputstring[25].ToString();
                            ExtractRow["KDPOS - Sales order item"] = inputstring[26].ToString();
                            ExtractRow["WWREN - CF invoice number"] = inputstring[27].ToString();
                            ExtractRow["CFCH00056 - Bill. Item"] = inputstring[28].ToString();
                            ExtractRow["WWBUN - BU"] = inputstring[29].ToString();
                            ExtractRow["WWPST - Product Structure"] = inputstring[30].ToString();
                            ExtractRow["WWIDS - Identstring"] = inputstring[31].ToString();
                            ExtractRow["WWPRS - Product segment"] = inputstring[32].ToString();
                            ExtractRow["WWHWK - Product characteristic"] = inputstring[33].ToString();
                            ExtractRow["WWFTT - Lacquering"] = inputstring[34].ToString();
                            ExtractRow["WWKAS- Coating"] = inputstring[35].ToString();
                            ExtractRow["WWDRU - Print"] = inputstring[36].ToString();
                            ExtractRow["WWEND - Final Form"] = inputstring[37].ToString();
                            ExtractRow["WWBRA - Brand"] = inputstring[38].ToString();
                            ExtractRow["MATKL - Material Group"] = inputstring[39].ToString();
                            ExtractRow["FRWAE - Local Currency"] = inputstring[40].ToString();
                            ExtractRow["MEINS - Sales Unit"] = inputstring[42].ToString();
                            ExtractRow["ABSMG - Sales quantity"] = inputstring[43].ToString();
                            ExtractRow["VV230 - Sales Volume KG"] = inputstring[44].ToString();
                            ExtractRow["VV998 - Periodic Quantity SQM"] = inputstring[45].ToString();





                            dtx.Rows.Add(ExtractRow);
                        }
                        else
                        {
                            if (ExtractRow[inputstring[6].ToString()].ToString() != "0")
                            {
                                AddCM Add1 = new AddCM();

                                try
                                {
                                    ExtractRow[inputstring[6].ToString()] = Add1.AddCMUtility(ExtractRow[inputstring[6].ToString()].ToString(), inputstring[41].ToString());
                                }
                                catch (Exception)
                                {
                                    MessageBox.Show(inputstring[6].ToString());
                                    throw;
                                }
                            }


                            //take # in column 41, period value, and post it in the corresponding contribution margin column in the new row
                            //MessageBox.Show(inputstring[6].ToString() + "---" + inputstring[41].ToString());
                            else
                                ExtractRow[inputstring[6].ToString()] = inputstring[41].ToString();

                        }
                    }
                }


                //dataGridView1.DataSource = dtx;
                dtx.ToCSV("C:/output/" + InputFileName.Substring(0, InputFileName.Length - 3) + "csv");
                dtx.ExportToExcel(InputFileName.Substring(0, InputFileName.Length - 3) + ".xlsx");
            }
        }

        public static DataTable OutputDataTable1()
        {
            DataTable dt = new DataTable();


            DataColumn OutputColumn;

            //BUDAT - Posting Date
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "BUDAT - Posting Date";
            dt.Columns.Add(OutputColumn);

            //FDAT - Invoice Date
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "FDAT - Invoice Date";
            dt.Columns.Add(OutputColumn);

            //BURKS - Company Code
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "BURKS - Company Code";
            dt.Columns.Add(OutputColumn);

            //VERSI - Version
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VERSI - Version";
            dt.Columns.Add(OutputColumn);

            //CPLYEAR - Planning Year
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CPLYEAR - Planning Year";
            dt.Columns.Add(OutputColumn);

            //0FISCYEAR-Planning Year
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "0FISCYEAR-Planning Year";
            dt.Columns.Add(OutputColumn);

            //WWPER - Fiscal Period
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPER - Fiscal Period";
            dt.Columns.Add(OutputColumn);

            //CCOMPANY - Company
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CCOMPANY - Company";
            dt.Columns.Add(OutputColumn);

            //VBUND - Partner Company
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VBUND - Partner Company";
            dt.Columns.Add(OutputColumn);

            //CFCH002620 - Item
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CFCH002620 - Item";
            dt.Columns.Add(OutputColumn);

            //KOKRS - Controlling Area
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KOKRS - Controlling Area";
            dt.Columns.Add(OutputColumn);

            //PRCTR - Profit Center
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "PRCTR - Profit Center";
            dt.Columns.Add(OutputColumn);


            //CPPRCTR - Partner Profit Center
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CPPRCTR - Partner Profit Center";
            dt.Columns.Add(OutputColumn);

            //WWBRN - Branch/Industry
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWBRN - Branch/Industry";
            dt.Columns.Add(OutputColumn);

            //WWPRG - Product Group
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPRG - Product Group";
            dt.Columns.Add(OutputColumn);


            //WWPPG - Partner Product Group
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPPG - Partner Product Group";
            dt.Columns.Add(OutputColumn);

            //WWART - Material Number
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWART - Material Number";
            dt.Columns.Add(OutputColumn);

            //KUNWE - Ship-To (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KUNWE - Ship-To (local)";
            dt.Columns.Add(OutputColumn);

            //KNDNR - Sold-To (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KNDNR - Sold-To (local)";
            dt.Columns.Add(OutputColumn);

            //KUNRE - Bill-To (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KUNRE - Bill-To (local)";
            dt.Columns.Add(OutputColumn);

            //KUNRG - Payer (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KUNRG - Payer (local)";
            dt.Columns.Add(OutputColumn);

            //WWKUN - Ship-To Final (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWKUN - Ship-To Final (local)";
            dt.Columns.Add(OutputColumn);

            //WWLWE - Country (Ship-To)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWLWE - Country (Ship-To)";
            dt.Columns.Add(OutputColumn);

            //LAND1 - Country (Sold-To)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "LAND1 - Country (Sold-To)";
            dt.Columns.Add(OutputColumn);

            //WWLRE - Country (Bill-To)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWLRE - Country (Bill-To)";
            dt.Columns.Add(OutputColumn);

            //WWLRG - Country (Payer)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWLRG - Country (Payer)";
            dt.Columns.Add(OutputColumn);

            //WWFCU - Country (Ship-To Final)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWFCU - Country (Ship-To Final)";
            dt.Columns.Add(OutputColumn);

            //KSTRG - Cost object
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KSTRG - Cost object";
            dt.Columns.Add(OutputColumn);

            //WWKAF - Sales order
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWKAF - Sales order";
            dt.Columns.Add(OutputColumn);

            //KDPOS - Sales order item
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KDPOS - Sales order item";
            dt.Columns.Add(OutputColumn);

            //WWREN - CF invoice number
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWREN - CF invoice number";
            dt.Columns.Add(OutputColumn);


            //CFCH00056 - Bill. Item
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CFCH00056 - Bill. Item";
            dt.Columns.Add(OutputColumn);

            //WWBUN - BU
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWBUN - BU";
            dt.Columns.Add(OutputColumn);

            //WWPST - Product Structure
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPST - Product Structure";
            dt.Columns.Add(OutputColumn);

            //WWIDS - Identstring
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWIDS - Identstring";
            dt.Columns.Add(OutputColumn);

            //WWPRS - Product segment
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPRS - Product segment";
            dt.Columns.Add(OutputColumn);

            //WWHWK - Product characteristic
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWHWK - Product characteristic";
            dt.Columns.Add(OutputColumn);

            //WWFTT - Lacquering
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWFTT - Lacquering";
            dt.Columns.Add(OutputColumn);

            //WWKAS- Coating
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWKAS- Coating";
            dt.Columns.Add(OutputColumn);

            //WWDRU - Print
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWDRU - Print";
            dt.Columns.Add(OutputColumn);

            //WWEND - Final Form
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWEND - Final Form";
            dt.Columns.Add(OutputColumn);

            //WWBRA - Brand
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWBRA - Brand";
            dt.Columns.Add(OutputColumn);

            //MATKL - Material Group
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "MATKL - Material Group";
            dt.Columns.Add(OutputColumn);

            //FRWAE - Local Currency
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "FRWAE - Local Currency";
            dt.Columns.Add(OutputColumn);

            //CM10010ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10010ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10015ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10015ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10020ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10020ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10040ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10040ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10050ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10050ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10070ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10070ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10080ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10080ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10090ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10090ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10100ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10100ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10130ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10130ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10140ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10140ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10150ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10150ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10170ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10170ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10180ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10180ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10190ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10190ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10200ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10200ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10210ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10210ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10250ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10250ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10260ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10260ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10270ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10270ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10280ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10280ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10290ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10290ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10300ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10300ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10330ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10330ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10340ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10340ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10350ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10350ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10360ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10360ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10370ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10370ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10380ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10380ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10410ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10410ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10420ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10420ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10430ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10430ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10440ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10440ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10450ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10450ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10460ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10460ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10490ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10490ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10500ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10500ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10510ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10510ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10520ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10520ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10530ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "FCM10530ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10540ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10540ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10550ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10550ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10560ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10560ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10570ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10570ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10580ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10580ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10610ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10610ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10620ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10620ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10630ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10630ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10650ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10650ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10660ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10660ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10670ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10670ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10680ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10680ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10700ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10700ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10710ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10710ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10720ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10720ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10730ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10730ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10740ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10740ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10760ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10760ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10770ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10770ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10810ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10810ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10820ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10820ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10830ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10830ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10840ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10840ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10850ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10850ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10860ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10860ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10870ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10870ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10910ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10910ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10920ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10920ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10930ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10930ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10940ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10940ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10950ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10950ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10960ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10960ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10980ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10980ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10990ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10990ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11000ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11000ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11020ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11020ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11030ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11030ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11040ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11040ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11050SUM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11050SUM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11070ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11070ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11080ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11080ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11090ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11090ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11100ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11100ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11120ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11120ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11125ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11125ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11140ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11140ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11145ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11145ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11160ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11160ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //MEINS - Sales Unit
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "MEINS - Sales Unit";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //ABSMG - Sales quantity
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "ABSMG - Sales quantity";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //VV230 - Sales Volume KG
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VV230 - Sales Volume KG";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //VV998 - Periodic Quantity SQM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VV998 - Periodic Quantity SQM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            return dt;





        }

        public string CompanyCode(string inputfilename)
        {

            if (inputfilename.Length == 0)
                return inputfilename;

            if (inputfilename.Substring(0, 4) == "GRAF")
            {
                //01234567891123456
                //GRAF_COPA_2017-10.txt
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1210";
            }
            if (inputfilename.Substring(0, 6) == "F-SUSM")
                return "1440";
            if (inputfilename.Substring(0, 6) == "F-SUSF")
                return "11450";
            if (inputfilename.Substring(0, 6) == "F-SUSC")
                return "1460";
            if (inputfilename.Substring(0, 6) == "F-SUSE")
                return "1470";
            if (inputfilename.Substring(0, 6) == "F-SPAF")
                return "1480";
            if (inputfilename.Substring(0, 6) == "F-SPGB")
                return "1490";
            if (inputfilename.Substring(0, 4) == "FACL")
            {
                //01234567891123456
                //FACL_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1690";

            }


            if (inputfilename.Substring(0, 4) == "FPEM")
            {
                //01234567891123456
                //FPEM_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1840";

            }


            if (inputfilename.Substring(0, 4) == "FPLV")
            {
                //01234567891123456
                //FPLV_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1850";

            }


            if (inputfilename.Substring(0, 4) == "FPLM")
            {
                //01234567891123456
                //FPLM_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1860";

            }

            if (inputfilename.Substring(0, 4) == "FPLI")
            {
                //01234567891123456
                //FPLI_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1870";

            }

            if (inputfilename.Substring(0, 4) == "FPLP")
            {
                //01234567891123456
                //FPLP_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "1880";

            }

            if (inputfilename.Substring(0, 6) == "H-HMUE")
                return "3000";
            if (inputfilename.Substring(0, 6) == "H-LABL")
                return "3020";
            if (inputfilename.Substring(0, 4) == "H-NO")
            {
                //01234567891123456
                //NOVI_COPA_2017-10
                
                return "3030";

            }



            if (inputfilename.Substring(0, 4) == "VERS")
            {
                //01234567891123456
                //VERS_COPA_2017-10
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3050";

            }


            if (inputfilename.Substring(0, 2) == "CM")
            {
                //0123456789112345
                //CM_COPA_2017-10
                fiscalmonth = inputfilename.Substring(13, 2);
                fiscalyear = inputfilename.Substring(8, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 2);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3060";
            }

            if (inputfilename.Substring(0, 4) == "CHIN")
            {
                //01234567891123456
                //CHIN_COPA_2017-10.txt
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3070";

            }
            if (inputfilename.Substring(0, 3) == "ETI")
            {
                //0123456789112345
                //ETI_COPA_2017-11
                fiscalmonth = inputfilename.Substring(14, 2);
                fiscalyear = inputfilename.Substring(9, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 3);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3120";
            }

            if (inputfilename.Substring(0, 2) == "EX")
            {
                //0123456789112345
                //EX_COPA_2017-11
                fiscalmonth = inputfilename.Substring(13, 2);
                fiscalyear = inputfilename.Substring(8, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 2);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3130";
            }
            if (inputfilename.Substring(0, 3) == "IMP")
            {
                //01234567891123456789212345
                //Upload4 - H-IMP - 2017.11.txt
                //UPLOAD4 - H-IMP 2017.10.txt
                fiscalmonth = inputfilename.Substring(23, 2);
                fiscalyear = inputfilename.Substring(18, 4);
                BUKRS = "H-" + inputfilename.Substring(12, 3);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3140";
            }
            if (inputfilename.Substring(0, 4) == "SLAB")
            {
                // 01234567891123456
                //SLAB_COPA_2017-10.txt
                fiscalmonth = inputfilename.Substring(15, 2);
                fiscalyear = inputfilename.Substring(10, 4);
                BUKRS = "H-" + inputfilename.Substring(0, 4);
                lastDayMonth = fiscalmonth + "/" + DateTime.DaysInMonth(Convert.ToInt32(fiscalyear), Convert.ToInt32(fiscalmonth)).ToString() + "/" + fiscalyear;
                return "3160";
            }

            MessageBox.Show("Didn't find Company Code");
            return inputfilename;


        }

        private void button4_Click(object sender, EventArgs e)
        {
            //streamReader to bring in files
            StreamReader streamreaderTXT;

            //creates an integer from the contribution margin for comparison to create new rows
            int outputItem = 0;
            int outputItemPrev = 0;

            //data row to be created to transpose CM scheme
            DataRow ExtractRow;

            //path of file
            string path = "";

            //identify "tab" as the column delimiter in the txt file
            char[] delimiter = new char[] { ',', ';', '\t' };

            //use the datatable created upon selection of files
            using (DataTable dtx = OutputDataTable1())
            {
                //loop through each file selected
                foreach (DataRow dt in fileArray.Rows)
                {

                    path = Path.GetFileNameWithoutExtension(dt[0].ToString());
                    streamreaderTXT = new StreamReader(dt[0].ToString());


                    ExtractRow = dtx.NewRow();

                    //loop through selected file
                    while (!streamreaderTXT.EndOfStream)
                    {
                        //inputstring is the entire row of data from the selected file
                        string[] inputstring = streamreaderTXT.ReadLine().Split(delimiter);

                        //if position not a contribution margin, skip
                        if (inputstring[6].Substring(0, 3) != "CM1")
                            continue;


                        outputItem = convertItem(inputstring[6]);


                        //if the current CM contribution margin scheme is less than or equal to, this indicates a new row should be created
                        //if outputItemPrev is 0, this indicates we are at the beginning of the file, and a new row should be created
                        if (outputItem <= outputItemPrev || outputItemPrev == 0)
                        {
                            //if this is not the first row, but a new row needs to be created, add the current data row to the overall data table
                            if (outputItemPrev != 0)
                                dtx.Rows.Add(ExtractRow);

                            //create a new row
                            ExtractRow = dtx.NewRow();


                            // insert into the data row,transposing the CM scheme of the current input row into a column in the data row, the inverse value
                            // of the value of the CM scheme 
                            //ExtractRow[inputstring[6].ToString()] = inverseString(inputstring[41].ToString());
                            ExtractRow[inputstring[6].ToString()] = inputstring[41].ToString();

                            //for creation of the data row, populate all other data besides the CM scheme values
                            ExtractRow["BUDAT - Posting Date"] = convertDate(inputstring[3].ToString(), inputstring[2].ToString());
                            ExtractRow["FDAT - Invoice Date"] = convertDate(inputstring[3].ToString(), inputstring[2].ToString());
                            ExtractRow["BURKS - Company Code"] = CompanyCode(inputstring[4].ToString());  //need to figure this piece out?
                            ExtractRow["VERSI - Version"] = inputstring[0].ToString();
                            ExtractRow["CPLYEAR - Planning Year"] = inputstring[1].ToString();
                            ExtractRow["0FISCYEAR-Planning Year"] = inputstring[2].ToString();
                            ExtractRow["WWPER - Fiscal Period"] = inputstring[3].ToString();
                            ExtractRow["CCOMPANY - Company"] = CompanyCode(inputstring[4].ToString());
                            ExtractRow["VBUND - Partner Company"] = CompanyCode(inputstring[5].ToString());
                            ExtractRow["CFCH002620 - Item"] = "";
                            ExtractRow["KOKRS - Controlling Area"] = inputstring[7].ToString();
                            ExtractRow["PRCTR - Profit Center"] = inputstring[8].ToString();
                            ExtractRow["CPPRCTR - Partner Profit Center"] = inputstring[9].ToString();
                            ExtractRow["WWBRN - Branch/Industry"] = inputstring[10].ToString();
                            ExtractRow["WWPRG - Product Group"] = inputstring[11].ToString();
                            ExtractRow["WWPPG - Partner Product Group"] = inputstring[12].ToString();
                            ExtractRow["WWART - Material Number"] = inputstring[13].ToString();//
                            ExtractRow["KUNWE - Ship-To (local)"] = inputstring[14].ToString();
                            ExtractRow["KNDNR - Sold-To (local)"] = inputstring[15].ToString();
                            ExtractRow["KUNRE - Bill-To (local)"] = inputstring[16].ToString();
                            ExtractRow["KUNRG - Payer (local)"] = inputstring[17].ToString();
                            ExtractRow["WWKUN - Ship-To Final (local)"] = inputstring[18].ToString();
                            ExtractRow["WWLWE - Country (Ship-To)"] = inputstring[19].ToString();
                            ExtractRow["LAND1 - Country (Sold-To)"] = inputstring[20].ToString();
                            ExtractRow["WWLRE - Country (Bill-To)"] = inputstring[21].ToString();
                            ExtractRow["WWLRG - Country (Payer)"] = inputstring[22].ToString();
                            ExtractRow["WWFCU - Country (Ship-To Final)"] = inputstring[23].ToString();
                            ExtractRow["KSTRG - Cost object"] = inputstring[24].ToString();
                            ExtractRow["WWKAF - Sales order"] = inputstring[25].ToString();
                            ExtractRow["KDPOS - Sales order item"] = inputstring[26].ToString();
                            ExtractRow["WWREN - CF invoice number"] = inputstring[27].ToString();
                            ExtractRow["CFCH00056 - Bill. Item"] = inputstring[28].ToString();//here
                            ExtractRow["WWBUN - BU"] = inputstring[29].ToString();
                            ExtractRow["WWPST - Product Structure"] = inputstring[30].ToString();
                            ExtractRow["WWIDS - Identstring"] = inputstring[31].ToString();
                            ExtractRow["WWPRS - Product segment"] = inputstring[32].ToString();
                            ExtractRow["WWHWK - Product characteristic"] = inputstring[33].ToString();
                            ExtractRow["WWFTT - Lacquering"] = inputstring[34].ToString();
                            ExtractRow["WWKAS- Coating"] = inputstring[35].ToString();
                            ExtractRow["WWDRU - Print"] = inputstring[36].ToString();
                            ExtractRow["WWEND - Final Form"] = inputstring[37].ToString();
                            ExtractRow["WWBRA - Brand"] = inputstring[38].ToString();
                            ExtractRow["MATKL - Material Group"] = inputstring[39].ToString();
                            ExtractRow["FRWAE - Local Currency"] = inputstring[40].ToString();
                            ExtractRow["MEINS - Sales Unit"] = inputstring[42].ToString();
                            ExtractRow["ABSMG - Sales quantity"] = inputstring[43].ToString();
                            ExtractRow["VV230 - Sales Volume KG"] = inputstring[44].ToString();
                            ExtractRow["VV998 - Periodic Quantity SQM"] = inputstring[45].ToString();

                        }//end if

                        else
                           // ExtractRow[inputstring[6].ToString()] = inverseString(inputstring[41].ToString());
                        ExtractRow[inputstring[6].ToString()] = inputstring[41].ToString();

                        //set output itemprev for comparision
                        outputItemPrev = outputItem;
                        //versionPrev = version;
                       // CCodePrev = CCode;
                        //counter++;

                    }//end while
                    dtx.Rows.Add(ExtractRow);

                }//end for each

                dtx.ToCSV("C:/output/" + path.ToString() + ".csv");
                dtx.Clear();
            }
        }

        public static int convertItem(string itemInput)
        {

            return Convert.ToInt16(itemInput.Substring(2, 5));


        }
        public static string inverseString(string input)
        {

            // if (input.Contains("105.38-"))
            // MessageBox.Show(input.ToString());
            if (input.Trim() == "")
                return "0";

            double i = 0;
            input = input.Replace("\"", "");

            if (input.Contains("-"))
            {

                //if "-" is at the end
                if (input.Substring(input.Length-1) == "-")
                {
                    i = Convert.ToDouble(input.Substring(0, input.Length - 1));
                    return i.ToString();
                }

                else
                {
                    i = Convert.ToDouble(input) * -1;
                    return i.ToString();
                }
            }

            i = Convert.ToDouble(input) * (-1);



            return i.ToString();

        }
        public static string convertDate(string month, string year)
        {
            /*if (month.Substring(1, 1) == "0")
                month = month.Substring(2, 1);
            else
                month = month.Substring(1, 2); */



            return month + "/" + DateTime.DaysInMonth(Convert.ToInt16(year), Convert.ToInt16(month)).ToString() + "/" + year;
        }
    }
}
