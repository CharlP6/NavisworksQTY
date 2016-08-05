using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet ExcelImport;

        #region ExcelImport

        private string GetConnectionString(string filePath)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Data Source"] = filePath;
            props["Extended Properties"] = "Excel 12.0 XML";


            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\MyExcel.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private DataSet ReadExcelFile(string filePath)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(filePath);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    foreach (DataRow dr in dtSheet.Rows)
                    {
                        string sheetName = dr["TABLE_NAME"].ToString();

                        cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                        cmd.Parameters.AddWithValue("TableName", sheetName);

                        DataTable dt = new DataTable();
                        dt.TableName = sheetName.Replace("\'", "").Replace("$", "");

                        try
                        {
                            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                            da.Fill(dt);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            lstStatus.Items.Add("Excel load failed: " + ex.Message);
                        }

                        ds.Tables.Add(dt);
                    }

                    cmd = null;
                    conn.Close();
                    lstStatus.Items.Add("Excel load success...");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    lstStatus.Items.Add("Excel load failed: " + ex.Message);
                }
            }
            return ds;
        }
        #endregion

        #region Variables

        List<DataTable> Tables = new List<DataTable>();
        List<DataRow> RawItemsData = new List<DataRow>();

        List<string> Hierarchy = new List<string>();
        List<string> RawItemsHeaders = new List<string>();
        List<string> SelectedHeaders = new List<string>();

        List<string[]> Processed = new List<string[]>();

        #endregion

        void InitializeImport()
        {
            foreach(DataTable dt in ExcelImport.Tables)
            {
                Tables.Add(dt);
            }

            if(Tables.Where(w => w.TableName == "Items Raw").Count() != 0)
            {

                lstStatus.Items.Add("\"Items Raw\" sheet found...");

                foreach(DataColumn dc in ExcelImport.Tables["Items Raw"].Columns)
                {
                    RawItemsHeaders.Add(dc.ColumnName);
                }
                
                int i = 1;

                while(i > 0)
                {
                    if(RawItemsHeaders.Contains("Group"+i.ToString()))
                    {
                        Hierarchy.Add("Group" + i.ToString());
                        i += 1;
                    }
                    else
                    {
                        i = 0;
                    }
                }

                if(Hierarchy.Count > 0)
                {
                    if (RawItemsHeaders.Contains("Item"))
                        Hierarchy.Add("Item");
                    else
                        lstStatus.Items.Add("No \"Item\" header found...");
                    

                    if (RawItemsHeaders.Contains("Object")) 
                        Hierarchy.Add("Object");
                    else
                        lstStatus.Items.Add("No \"Object\" header found...");

                    foreach(DataRow dr in ExcelImport.Tables["Items Raw"].Rows)
                    {
                        RawItemsData.Add(dr);
                    }
                    lstStatus.Items.Add("Hierarchy load successful...");
                }
                else
                {
                    lstStatus.Items.Add("No \"Group\" headers found. Please make sure the file exported correctly...");
                    MessageBox.Show("No headers found to group data. Please make sure that the file exported correctly.");
                }

                if (RawItemsHeaders.Contains("PrimaryQuantity")) SelectedHeaders.Add("PrimaryQuantity");
                else
                {
                    MessageBox.Show("Could not find \"PrimaryQuantity\" header.");
                }

                if (RawItemsHeaders.Contains("PrimaryQuantity Units")) SelectedHeaders.Add("PrimaryQuantity Units");
                else
                {
                    MessageBox.Show("Could not find \"PrimaryQuantity Units\" header.");
                }

                if (RawItemsHeaders.Contains("WBS")) SelectedHeaders.Add("WBS");
                else
                {
                    MessageBox.Show("Could not find \"WBS\" header.");
                }

            }
            else
            {
                MessageBox.Show("Sheet: \"Raw Items\" not found in loaded spreadsheet.\nPlease make sure that the correct file is loaded and that the data exported correctly.");
                lstStatus.Items.Add("\"Items Raw\" sheet not found...");
            }
        }

        void UpdateUI()
        {

            lstGroups.DataSource = null;
            lstHeaders.DataSource = null;
            lstColumns.DataSource = null;

            lstGroups.DataSource = Hierarchy;
            lstHeaders.DataSource = SelectedHeaders;
            lstColumns.DataSource = RawItemsHeaders.Where(w => !SelectedHeaders.Contains(w) && !Hierarchy.Contains(w)).ToList();

            lvDataView.Columns.Clear();

            lvDataView.Columns.Add("Level", 50);
            lvDataView.Columns.Add("Bill Item", 250);

            foreach(string s in SelectedHeaders)
            {                
                lvDataView.Columns.Add(s, TextRenderer.MeasureText(s, this.Font).Width+5);
            }
        }
        
        void TraverseData()
        {
            #region Initialize Iteration Variables
            int Level = 0; //Indicates current level of hierarchy iteration
            int maxLevel = Hierarchy.Count - 1; //Highest level that can be iterated
            int MaxSize = Hierarchy.Count; //Maximum array size based on hierarchy levels

            Queue<string>[] DataTree = new Queue<string>[MaxSize]; //Array of Queue objects, one for each hierarchy level

            for (int i = 0; i <= maxLevel; i++)
                DataTree[i] = new Queue<string>(); //initialize each queue in the array

            string[] LevelString = new string[MaxSize]; //String array saving the current hierarchy steps

            List<string> Group = Hierarchy; //String list for names of hierarchy levels
            #endregion

            RawItemsData.Select(s => s[Group[Level]] == DBNull.Value ? "" : (string)s[Group[Level]]).Distinct().ToList().ForEach(f => DataTree[Level].Enqueue(f)); //Populate first stack with level 0 data
            

            while(Level >= 0) //Ieterate as long as data is on the level 0 stack
            {
                if (DataTree[Level].Count > 0)
                {

                    LevelString[Level] = DataTree[Level].Dequeue();

                    IEnumerable<DataRow> StackEnum = RawItemsData.Where(w => ((w[Group[0]] == DBNull.Value) ? "" : (string)w[Group[0]]) == LevelString[0]); //Initialize enumerator

                    for (int i = 1; i <= Level; i++)
                    {
                        string ls = LevelString[i];
                        string gs = Group[i];

                        StackEnum = StackEnum.Where(w => (w[gs] == DBNull.Value ? "" : (string)w[gs]) == ls); //Filter enumerator for each stack leve
                    }

                    if (Level < maxLevel)
                    {
                        List<string> tList = StackEnum.Select(s => s[Group[Level + 1]] == DBNull.Value ? "" : (string)s[Group[Level + 1]]).Distinct().ToList();
                        tList.ForEach(f => DataTree[Level + 1].Enqueue(f)); //Fill next stack with current filter

                        if (LevelString[Level] != "")
                        {
                            int lvl = LevelString.Take(Level + 1).Where(w => w == "").Count();

                            string[] DisplayLine = new string[lvDataView.Columns.Count];

                            DisplayLine[0] = (Level - lvl+1).ToString();
                            DisplayLine[1] = LevelString[Level];

                            if(Level <= maxLevel - 2)
                            {
                                for (int i = 2; i < lvDataView.Columns.Count; i++)
                                    DisplayLine[i] = "";

                                ListViewItem lvi = new ListViewItem(DisplayLine);
                                lvDataView.Items.Add(lvi);
                                Processed.Add(DisplayLine);
                            }
                            else if(Level == maxLevel - 1)
                            {
                                DisplayLine[2] = StackEnum.Sum(su => su["PrimaryQuantity"] == DBNull.Value ? 0 : (double)su["PrimaryQuantity"]).ToString();
                                DisplayLine[3] = StackEnum.Select(s => s["PrimaryQuantity Units"] == DBNull.Value ? "" : (string)s["PrimaryQuantity Units"]).First();
                                for (int i = 4; i < lvDataView.Columns.Count; i++)
                                {
                                    string cS = lvDataView.Columns[i].Text;
                                    DisplayLine[i] = StackEnum.Select(s => s[cS] == DBNull.Value ? "" : s[cS].ToString()).First();
                                }
                                ListViewItem lvi = new ListViewItem(DisplayLine);
                                lvDataView.Items.Add(lvi);
                                Processed.Add(DisplayLine);
                            }
                        }
                    }
                    else
                    {
                        if (LevelString[Level] != "")
                        {
                            int lvl = LevelString.Take(Level + 1).Where(w => w == "").Count(); //Count number of blank levels to shift levels up accordingly;
                            string[] DisplayLine = new string[lvDataView.Columns.Count];

                            DisplayLine[0] = (Level - lvl+1).ToString();
                            DisplayLine[1] = LevelString[Level];
                            DisplayLine[2] = StackEnum.Sum(su => su["PrimaryQuantity"] == DBNull.Value ? 0 : (double)su["PrimaryQuantity"]).ToString();
                            DisplayLine[3] = StackEnum.Select(s => s["PrimaryQuantity Units"] == DBNull.Value ? "" : (string)s["PrimaryQuantity Units"]).First();
                            for (int i = 4; i < lvDataView.Columns.Count; i++)
                            {
                                string cS = lvDataView.Columns[i].Text;
                                DisplayLine[i] = StackEnum.Select(s => s[cS] == DBNull.Value ? "" : s[cS].ToString()).First();
                            }
                            ListViewItem lvi = new ListViewItem(DisplayLine);
                            lvDataView.Items.Add(lvi);
                            Processed.Add(DisplayLine);
                        }
                    }
                    Level = (Level + 1) > maxLevel ? maxLevel : Level + 1;
                }
                else
                {
                    Level -= 1;
                }
            }
        }

        private void loadQTY_Click(object sender, EventArgs e)
        {
            Hierarchy.Clear();
            RawItemsData.Clear();
            RawItemsHeaders.Clear();
            SelectedHeaders.Clear();
            Processed.Clear();

            DialogResult dr = openFileDialog1.ShowDialog();
            if(dr == System.Windows.Forms.DialogResult.OK)
            {
                ExcelImport = ReadExcelFile(openFileDialog1.FileName);

                if (ExcelImport != null)
                {
                    InitializeImport();
                    UpdateUI();
                    lvDataView.Items.Clear();
                    TraverseData();

                }
                else
                {
                    MessageBox.Show("Load Unsuccessful.");
                }
            }

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btnAddHeader_Click(object sender, EventArgs e)
        {
            if(lstColumns.SelectedItem != null)
            {
                SelectedHeaders.Add((string)lstColumns.SelectedItem);
            }
            UpdateUI();
            if (ExcelImport != null)
            {
                lvDataView.Items.Clear();
                Processed.Clear();
                TraverseData();
            }
        }

        private void btnRemoveHeader_Click(object sender, EventArgs e)
        {
            if(lstHeaders.SelectedItem != null)
            {
                string s = (string)lstHeaders.SelectedItem;
                if(s != "PrimaryQuantity" && s != "PrimaryQuantity Units")
                SelectedHeaders.Remove(s);

            }
            UpdateUI();
            if (ExcelImport != null)
            {
                lvDataView.Items.Clear();
                Processed.Clear();
                TraverseData();
            }
        }

        private void exportCCSSheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //WriteExcelFile();
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (ExcelImport != null)
            {
                if (dr == DialogResult.OK)
                {
                    SaveExcel();
                }
            }
            else
            {
                MessageBox.Show("No data found to export...");
            }
        }

        private void WriteExcelFile()
        {
            List<string> h = new List<string>();
            foreach(ColumnHeader ch in lvDataView.Columns)
            {
                h.Add(ch.Text);
            }
            string headernames = string.Join(", ", h.Select(s => s.Replace(",", ".")).ToArray());
            
            DialogResult dr = saveFileDialog1.ShowDialog();

            if (dr == DialogResult.OK)
            {
                using (FileStream fs = new FileStream(saveFileDialog1.FileName, FileMode.Create, FileAccess.Write))
                {
                    using(StreamWriter sw = new StreamWriter(fs,Encoding.UTF8))
                    {
                        sw.WriteLine(headernames);
                        foreach(string[] s in Processed)
                        {
                            sw.WriteLine(string.Join(",", s.Select(se => se.Replace(",", ".")).ToArray()));
                        }
                    }
                }
            }
        }

        void SaveExcel()
        {
            string connectionString = GetConnectionString(saveFileDialog1.FileName);

            List<string> h = new List<string>();
            foreach (ColumnHeader ch in lvDataView.Columns)
            {
                h.Add(ch.Text);
            }

            string ColumnDeclaration = "([" + string.Join("] VARCHAR, [", h.ToArray()) + "] VARCHAR)";
            string ColumnNames = "([" + string.Join("],[", h.ToArray()) + "])";

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = "CREATE TABLE [CandyExport] " + ColumnDeclaration;
                cmd.ExecuteNonQuery();

                foreach(string[] s in Processed)
                {
                    string Written = string.Join(",", s.Select(se => "?").ToArray());
                    cmd.CommandText = "INSERT INTO [CandyExport]" + ColumnNames + " VALUES(" + Written + ")";

                    s.Select(se => se).ToList().ForEach(f => cmd.Parameters.AddWithValue("?", f));

                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }



                conn.Close();
            }
        }

    }
}
