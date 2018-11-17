using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private static readonly log4net.ILog log =
       log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = Filepath.Text.Trim();
                ProcessDirectory(filepath);
            }
            catch (Exception ex)
            {

                log.Error("Error Message:" + ex.Message);
            }
        }

        public static void ProcessDirectory(string targetDirectory)
        {
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (subdirectoryEntries.Length == 0)
            {
                string[] fileEntries = Directory.GetFiles(targetDirectory);
                foreach (string fileName in fileEntries)
                    ProcessFileAsync(fileName);
                    MessageBox.Show("Files Imported Successfully");
            }
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory);
            }

        }

        public static async Task ProcessFileAsync(string path)
        {

            try
            {
                string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes;IMEX=1'", path);

                using (OleDbConnection con = new OleDbConnection(connString))
                {
                    con.Open();
                    var sheetNames = con.GetSchema("Tables");
                    string row1 = sheetNames.Rows[0]["TABLE_NAME"].ToString();
                    //foreach (DataRow row in sheetNames.Rows)
                    //{
                    //    var name = row["TABLE_NAME"];
                    //    string name1 = name.ToString().Replace('"', ' ').Replace("'", string.Empty);
                    System.Data.DataTable dt = GetDataTable("SELECT * FROM " + "[" + row1 + "]", connString, con, row1);
                    System.Data.DataTable newdt = RemoveEmptyRowsFromDataTable(dt);
                    int i = 0;
                    log.Error("Start Date:" + DateTime.Now + "-" + "File Path:" + path);
                    Int64 rows = newdt.Rows.Count;
                    //for (int j = 0; j < newdt.Rows.Count; j++)
                    //{
                    //    ed.TestName = newdt.Columns.Contains("Test Name") ? newdt.Rows[j]["Test Name"].ToString() : null;
                    //    ed.Inventory = newdt.Columns.Contains("Inventory") ? newdt.Rows[j]["Inventory"].ToString() : null;
                    //    list.Add(ed);
                    //}

                    foreach (DataRow row in dt.Rows)
                    {
                        EntityData ed = new EntityData();
                        ed.TestName = row["Test Name"].ToString();
                        ed.Inventory = row["Inventory"].ToString();
                        list.Add(ed);
                    }

                    HttpClient client = new HttpClient();
                    Uri baseAddress = new Uri("http://localhost:15751/");
                    client.BaseAddress = baseAddress;
                    HttpResponseMessage response = client.PostAsJsonAsync("api/MasterTest/MasterTestInventoryRelationBulk", list).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        //MessageBox.Show("Record Imported Successfully");
                    }
                    //}

                    //}
                    log.Error("End Date:" + DateTime.Now);
                }
            }
            catch (Exception ex)
            {
                log.Error("Error Message:" + ex.Message);
            }
        }
    }
}
