using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
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

        #region Excel

        private string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = @"D:\Users\pretoriusc\Desktop\Project Temp Folder\Navisworks QTY\NW qty Export.xlsx";

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

        private DataSet ReadExcelFile()
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString();

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
                        }

                        ds.Tables.Add(dt);
                    }

                    cmd = null;
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);                
                }
            }
            return ds;
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelImport = ReadExcelFile();

            foreach (DataTable dt in ExcelImport.Tables)
            {
                lstTables.Items.Add(dt.TableName);
            }

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstColumns.Items.Clear();


            foreach (DataColumn dc in ExcelImport.Tables[lstTables.SelectedItem.ToString()].Columns)
            {
                lstColumns.Items.Add(dc.ColumnName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            lstGroups.Items.Add(lstColumns.SelectedItem.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox6.Items.Clear();

            //create a real list of data rows
            List<DataRow> TableRows = new List<DataRow>();
            foreach(DataRow dr in ExcelImport.Tables[lstTables.SelectedItem.ToString()].Rows)
            {
                TableRows.Add(dr);
            }

            #region Initialize Iteration Variables
            int Level = 0;
            int maxLevel = lstGroups.Items.Count - 1;
            int MaxSize = lstGroups.Items.Count;

            Queue<string>[] DataTree = new Queue<string>[MaxSize];
            for (int i = 0; i <= maxLevel; i++)
                DataTree[i] = new Queue<string>();

            string[] LevelString = new string[MaxSize];

            string[] Group = new string[MaxSize];

            for (int i = 0; i < MaxSize; i++)
                Group[i] = lstGroups.Items[i].ToString();
            #endregion

            TableRows.Select(s => s[Group[Level]] == DBNull.Value ? "" : (string)s[Group[Level]]).Distinct().ToList().ForEach(f => DataTree[Level].Enqueue(f)); //Populate first stack
            

            while(Level >= 0)
            {
                if (DataTree[Level].Count > 0)
                {

                    LevelString[Level] = DataTree[Level].Dequeue();

                    IEnumerable<DataRow> StackEnum = TableRows.Where(w => ((w[Group[0]] == DBNull.Value) ? "" : (string)w[Group[0]]) == LevelString[0]);

                    for (int i = 1; i <= Level; i++)
                    {
                        string ls = LevelString[i];
                        string gs = Group[i];

                        StackEnum = StackEnum.Where(w => (w[gs] == DBNull.Value ? "" : (string)w[gs]) == ls);


                    }

                    if (Level < maxLevel)
                    {
                        List<string> tList = StackEnum.Select(s => s[Group[Level + 1]] == DBNull.Value ? "" : (string)s[Group[Level + 1]]).Distinct().ToList();
                        tList.ForEach(f => DataTree[Level + 1].Enqueue(f));

                        if (LevelString[Level] != "")
                        {
                            int lvl = LevelString.Take(Level + 1).Where(w => w == "").Count();
                            listBox6.Items.Add((Level - lvl).ToString() + ": " + LevelString[Level]);
                        }
                    }
                    else
                    {
                        int lvl = 0;
                        if (LevelString[Level] != "")
                        {
                            lvl = LevelString.Take(Level + 1).Where(w => w == "").Count();
                        }

                        listBox6.Items.Add((Level - lvl).ToString() + ": " + LevelString[Level] + "\t" + StackEnum.Sum(su => su["PrimaryQuantity"] == DBNull.Value ? 0 : (double)su["PrimaryQuantity"]).ToString());
                    }



                    Level = (Level + 1) > maxLevel ? maxLevel : Level + 1;

                }
                else
                {
                    Level -= 1;
                }
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            lstGroups.Items.RemoveAt(lstGroups.SelectedIndex);
        }

    }

    class Hierarchy
    {
        public int Level { get; set; }
        public string Name { get; set; }
    }
}
