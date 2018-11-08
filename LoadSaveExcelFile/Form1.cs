using System;

using System.Data;
using System.Data.OleDb;

using System.IO;

using System.Windows.Forms;

namespace LoadSaveExcelFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private OleDbConnection conn;
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "(.xls)|*.xls|(.xlsx)|*.xlsx|(.xlsm)|*.xlsm";
            opf.ShowDialog();

            string path = opf.FileName;

            string extintion = Path.GetExtension(path);
            string pathcon = "";
            textBox1.Text = path;
            if (extintion == ".xls")
            {
                pathcon =
                 @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            }
            else if (extintion == ".xlsx" || extintion == ".xlsm")
            {
                pathcon = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;IMEX=2.0;HDR=YES""", path);



            }





            conn = new OleDbConnection(pathcon);
            conn.Open();
            DataTable sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (DataRow sheet in sheets.Rows)
            {
                string sht = sheet[2].ToString().Replace("0", "0");

                comboBoxsheet.Items.Add(sht);
            }

            conn.Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            conn.Open();
            string selectitem = comboBoxsheet.SelectedItem.ToString();

            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("select * from [" + selectitem + "]", conn);


            DataTable dt = new DataTable();

            myDataAdapter.Fill(dt);

            gridControl1.DataSource = dt;
            conn.Close();
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog opf = new FolderBrowserDialog();
            //opf.ShowDialog();


            //string path = opf.SelectedPath;
            //string filename = textBox2.Text;
          //  string selectextion = comboextntion.SelectedItem.ToString();
            gridControl1.ShowPrintPreview();
           // gridControl1.ExportToXls(Path.Combine(path, "" + filename + "" + selectextion + ""));

           // MessageBox.Show("Excel File  " + filename + " Save");
        }
    }
}

