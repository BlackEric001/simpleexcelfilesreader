using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;

namespace ExcelReadTest
{
    public delegate void writeLog(string logMessage);

    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            this.Text = APP_NAME;
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";

            listBox1.ContextMenuStrip.Enabled = false;

            excelReader.log = new writeLog(this.WriteLog);
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                readFile(openFileDialog1.FileName);
                this.Text = APP_NAME + "  " + openFileDialog1.FileName;
            }
        }

        private void readFile(string fileName)
        {
            DataTable oTbl = new DataTable();

            excelReader.excelFileName = fileName;
            
            try
            {
                if (!excelReader.getTables(ref oTbl))
                    WriteLog("Error get tables");
                else
                {
                    tableNames.Clear();

                    foreach (DataRow row in oTbl.Rows)
                    {
                        string data = String.Empty;
                        foreach (DataColumn dc in oTbl.Columns)
                        {
                            var field1 = row[dc].ToString();
                            data += "[ " + dc.Caption + ": " + field1 + "]";
                            if (dc.Caption == TABLE_NAME_COLUMN)
                                tableNames.Add(field1.ToString());

                        }
                        WriteLog(data);
                    }
                    displayTablesList();
                }

                if (tableNames.Count > 0)
                    loadExcelListData(tableNames[0]);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            finally
            {
                if (oTbl != null) oTbl.Dispose();
            }
        }

        private void cbHDR_CheckedChanged(object sender, EventArgs e)
        {
            if (cbHDR.Checked)
                excelReader.HDR = "HDR=YES;";
            else
                excelReader.HDR = "HDR=No;";
        }

        private void displayTablesList()
        {
            listBox1.Items.Clear();

            foreach (string tName in tableNames)
                listBox1.Items.Add(tName);

            listBox1.ContextMenuStrip.Enabled = listBox1.Items.Count > 0;
        }



        private void LoadListData_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
                loadExcelListData(listBox1.SelectedItem.ToString());
            else
                MessageBox.Show("Выберите лист для загрузки данных!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void loadExcelListData(string listName)
        {
            WriteLog("Try load data from " + listName);
            DataTable dt = new DataTable();
            try
            {
                if (!excelReader.getListData(listName, ref dt))
                    WriteLog("Невозможно получить данные с листа " + listName);
                else
                    dataGridView1.DataSource = dt;
            }
            finally
            {
                if (dt != null) dt.Dispose();
            }

        }

        private void WriteLog(string message)
        {
            richTextBox1.AppendText(Environment.NewLine + DateTime.Now + "   " + message);
        }


        ExcelReader excelReader = new ExcelReader();

        private List<string> tableNames = new List<string>();

        private const string TABLE_NAME_COLUMN = "TABLE_NAME";
        private const string APP_NAME = "Excel File Reader";
       
    }
}
