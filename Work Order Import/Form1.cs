using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Data.SQLite;

namespace Work_Order_Import
{
    public partial class Form1 : Form
    {
        private int rowIndex;
        //Sets SQL server connection string based on which location employee is located at
        public static string GetColumbusString(string columbusLocation)
        {
            string connectionString;
            switch (columbusLocation)
            {
                case "Eugene":
                    connectionString = @"Server = Eugene2\SQLEXPRESS; Database = Ideal Steel; Trusted_Connection = True; ";
                    return connectionString;
                case "Seneca":
                    connectionString = @"Server = Seneca2\SQLEXPRESS; Database = IdealSteel-Seneca 05 2014; Trusted_Connection = True;";
                    return connectionString;
                case "Springfield":
                    connectionString = @"Server = ID-OH-006\SQLEXPRESS; Database = Ohio Columbus 2019; Trusted_Connection = True;";
                    return connectionString;
                case "Houston":
                    connectionString = @"Server = GQYPK02\ESAB; Database = IdealSteel_Houston_Columbus_5_2014; Trusted_Connection = True;";
                    return connectionString;
                case "Demo":
                    string dbPath = Properties.Settings.Default.demoDbPath;
                    if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
                    {
                        MessageBox.Show("Demo database path is not set or invalid.");
                        return null;
                    }
                    return $"Data Source={dbPath};Version=3;";
                default:
                    connectionString = "null";
                    return connectionString;
            }
        }
        //Returns Columbus SQL query
        public static string ColumbusQuery()
        {
            return @"SELECT
            UPPER(Parts.Name) AS PartName,
            ROUND(Thicknesses.Thickness / 25.4, 3) AS Thickness,
            UPPER(Materials.Name) AS Material,
            ROUND(((Parts.Area / 92903.04) * 144 * (Thicknesses.Thickness / 25.4)) * 0.2836, 0) AS Weight,
            PartObjectTS AS DateImported
            FROM Parts
            INNER JOIN Materials ON Parts.ID_MaterialDefault = Materials.ID
            INNER JOIN Thicknesses ON Parts.ID_ThicknessDefault = Thicknesses.ID
            ORDER BY Parts.Name ASC;";
        }
        //Returns MAS ERP SQL query
        public static string MasQuery()
        {
            return @"SELECT ItemCode, QuantityOrdered, PromiseDate, D404_WorkOrderNo
            FROM SO_SalesOrderDetail
            WHERE D404_WorkOrderNo IS NOT NULL";
        }
        //Checks if demo database is selected, otherwise returns connection string for MAS
        public static string GetMasString()
        {
            string location = Properties.Settings.Default.nestingLocation;

            if (location == "Demo")
            {
                string dbPath = Properties.Settings.Default.demoDbPath;

                if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
                {
                    MessageBox.Show("Demo MAS database path is not set or invalid.");
                    return null;
                }

                return $"Data Source={dbPath};Version=3;";
            }

            return @"Password=P@$$w0rd; User ID=syncserver; Initial Catalog=WOScanServer; Server=mas\WOSCAN2;";
        }
        //Method to populate Columbus database info into dataGridView3
        public void PopulateColumbusData()
        {
            string location = Properties.Settings.Default.nestingLocation;
            string connectionString = GetColumbusString(location);
            string query = ColumbusQuery();

            try
            {
                if (location == "Demo")
                {
                    using (var conn = new SQLiteConnection(connectionString))
                    {
                        conn.Open();
                        var cmd = new SQLiteCommand(query, conn);
                        var adapter = new SQLiteDataAdapter(cmd);
                        var ds = new DataSet();
                        adapter.Fill(ds);
                        dataGridView3.ReadOnly = false;
                        dataGridView3.DataSource = ds.Tables[0];
                        conn.Close();
                    }
                }
                else
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        var cmd = new SqlCommand(query, conn);
                        var adapter = new SqlDataAdapter(cmd);
                        var ds = new DataSet();
                        adapter.Fill(ds);
                        dataGridView3.ReadOnly = false;
                        dataGridView3.DataSource = ds.Tables[0];
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Columbus data: " + ex.Message);
            }
        }    
        //Method to populate MAS database info into dataGridView4
        public void PopulateMasData()
        {
            string location = Properties.Settings.Default.nestingLocation;
            string connectionString = GetMasString();
            string query = MasQuery();

            try
            {
                if (location == "Demo")
                {
                    using (var conn = new SQLiteConnection(connectionString))
                    {
                        conn.Open();
                        var cmd = new SQLiteCommand(query, conn);
                        var adapter = new SQLiteDataAdapter(cmd);
                        var ds = new DataSet();
                        adapter.Fill(ds);
                        dataGridView4.ReadOnly = false;
                        dataGridView4.DataSource = ds.Tables[0];
                        conn.Close();
                    }
                }
                else
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        var cmd = new SqlCommand(query, conn);
                        var adapter = new SqlDataAdapter(cmd);
                        var ds = new DataSet();
                        adapter.Fill(ds);
                        dataGridView4.ReadOnly = false;
                        dataGridView4.DataSource = ds.Tables[0];
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading MAS data: " + ex.Message);
            }
        }
        //Used for part matching when part number in Columbus does not match MAS part number and reconciles this using part match file        
        private void InsertDgvIntoForm()
        {
            //Ensure fileLocation is set
            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.fileLocation) || !Directory.Exists(Properties.Settings.Default.fileLocation))
            {
                MessageBox.Show("Please select a location to store your part match file (User Settings.xml).");

                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select folder for part match file";

                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        Properties.Settings.Default.fileLocation = folderDialog.SelectedPath;
                        Properties.Settings.Default.Save();
                    }
                    else
                    {
                        MessageBox.Show("A valid location is required to run the program.");
                        Application.Exit();
                        return;
                    }
                }
            }
            //Ensure XML file exists
            string xmlPath = Path.Combine(Properties.Settings.Default.fileLocation, "User Settings.xml");

            if (!File.Exists(xmlPath))
            {
                MessageBox.Show("Part number matching file not found. Creating a new one now.");

                //Create a default table and write it to XML
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                dt.Columns.Add("W/O Name");
                dt.Columns.Add("Columbus Name");
                dt.Rows.Add("", "");
                ds.Tables.Add(dt);
                ds.WriteXml(xmlPath);
            }

            //Load the XML file into dataGridView2
            try
            {
                XmlReader xmlFile = XmlReader.Create(xmlPath, new XmlReaderSettings());
                DataSet ds = new DataSet();
                ds.ReadXml(xmlFile);
                xmlFile.Close();

                dataGridView2.DataSource = ds.Tables[0];
                dataGridView2.Columns[0].Width = 165;
                dataGridView2.Columns[1].Width = 160;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading part match file: " + ex.Message);
            }
        }
        //Method for taking dataGridView2 part number matching and exporting it out to an XML file for repeated use
        private void ExportDgvToXML()
        {
            if (File.Exists($"{Properties.Settings.Default.fileLocation}\\User Settings.xml"))
            {
                File.Delete($"{Properties.Settings.Default.fileLocation}\\User Settings.xml");
            }
            DataTable dt = (DataTable)dataGridView2.DataSource;
            dt.WriteXml($"{Properties.Settings.Default.fileLocation}\\User Settings.xml");
        }
        //Checks if work order exists in dataGridView4
        private void CheckifWOExists(string woNumber, out bool test)
        {
            test = false;

            try
            {
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    if (row.IsNewRow) continue;
                    var cell = row.Cells[3];
                    if (cell?.Value != null && cell.Value.ToString().Contains(woNumber))
                    {
                        test = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in CheckifWOExists: " + ex.Message);
            }
        }
        //Pulls work order info from dataGridView4         
        private void GetWorkOrderInfo(string woNumber, out string partNumber, out string orderAmount, out string dueDate)
        {
            partNumber = orderAmount = dueDate = string.Empty;

            try
            {
                DataGridViewRow row = dataGridView4.Rows
                    .Cast<DataGridViewRow>()
                    .FirstOrDefault(r => !r.IsNewRow && r.Cells[3].Value != null && r.Cells[3].Value.ToString().Equals(woNumber));

                if (row == null)
                {
                    MessageBox.Show("Work order not found in MAS data.");
                    return;
                }

                partNumber = row.Cells[0].Value?.ToString().ToUpper() ?? "";
                orderAmount = row.Cells[1].Value?.ToString() ?? "";
                dueDate = row.Cells[2].Value?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in GetWorkOrderInfo: " + ex.Message);
            }
        }
        //Inputs work order info into dataGridView1 to be imported into columbus
        private void InputWOInfo()
        {
            string woNumber = textBox1.Text.ToString();
            CheckifWOExists(woNumber, out bool test);
            if (test == false)
            {
                //Allows for manual entry of work order info if not found in table
                MessageBox.Show("Work order number could not be found!");
                DialogResult dialogResult = MessageBox.Show("Would you like to enter work order manually?", "", MessageBoxButtons.YesNo);
                if(dialogResult == DialogResult.Yes)
                {
                    string mwoNumber = Microsoft.VisualBasic.Interaction.InputBox("Please enter work order number:",
                        "User Input",
                        "",
                        200,
                        200);
                    string mpartNumber = Microsoft.VisualBasic.Interaction.InputBox("Please enter part number:",
                        "User Input",
                        "",
                        200,
                        200);
                    string mThickness = Microsoft.VisualBasic.Interaction.InputBox("Please enter thickness:",
                        "User Input",
                        "",
                        200,
                        200);
                    string mMaterial = Microsoft.VisualBasic.Interaction.InputBox("Please enter material:",
                        "User Input",
                        "",
                        200,
                        200);
                    string mQuantity = Microsoft.VisualBasic.Interaction.InputBox("Please enter order quantity:",
                        "User Input",
                        "",
                        200,
                        200);
                    string mdueDate = Microsoft.VisualBasic.Interaction.InputBox("Please enter due date:",
                        "User Input",
                        "",
                        200,
                        200);
                    if (dataGridView1.Rows[0].Cells[1].Value == null)
                    {
                        dataGridView1.Rows[0].SetValues(mwoNumber, mpartNumber, mThickness, mMaterial, mQuantity, false, mdueDate, " ");
                    }
                    else 
                    { 
                        dataGridView1.Rows.Add(mwoNumber, mpartNumber, mThickness, mMaterial, mQuantity, false, mdueDate, " "); 
                    }
                    int tRowIndex = dataGridView1.Rows.Count - 1;
                    int tColumnIndex = 5;
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[tRowIndex].Selected = true;
                    dataGridView1.Rows[tRowIndex].Cells[tColumnIndex].Selected = true;
                    dataGridView1.FirstDisplayedScrollingRowIndex = tRowIndex;
                    textBox1.Select();
                    textBox1.Clear();
                    return;
                }
                else
                return;
            }
            GetWorkOrderInfo(woNumber, out string partNumber, out string orderAmount, out string dueDate);
            CheckIfPNExists(partNumber, out bool test2, out string partNumberFinal);
            if (decimal.TryParse(orderAmount, out decimal parsedQty))
            {
                orderAmount = ((int)parsedQty).ToString();
            }
            else
            {
                MessageBox.Show("Invalid order quantity format.");
                return;
            }
            if (DateTime.TryParse(dueDate, out DateTime parsedDate))
            {
                dueDate = parsedDate.ToString("yyyy-MM-dd");
            }
            else
            {
                MessageBox.Show("Invalid due date format.");
                return;
            }
            if (test2 == false)
            {
                return;
            }
            GetColumbusData(partNumberFinal, out string materialGrade, out string thickness);
            if (dataGridView1.Rows[0].Cells[1].Value == null)
            {
                dataGridView1.Rows[0].SetValues(woNumber, partNumberFinal, thickness, materialGrade, orderAmount, false, dueDate, " ");
            }
            else { dataGridView1.Rows.Add(woNumber, partNumberFinal, thickness, materialGrade, orderAmount, false, dueDate, " "); }
            int nRowIndex = dataGridView1.Rows.Count - 1;
            int nColumnIndex = 5;
            dataGridView1.ClearSelection();
            dataGridView1.Rows[nRowIndex].Selected = true;
            dataGridView1.Rows[nRowIndex].Cells[nColumnIndex].Selected = true;
            dataGridView1.FirstDisplayedScrollingRowIndex = nRowIndex;
            textBox1.Select();
            textBox1.Clear();
        }
        //Checks if part number exists in columbus database
        private void CheckIfPNExists(string partNumber, out bool test2, out string partNumberFinal)
        {
            test2 = false;
            partNumberFinal = null;

            try
            {
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.IsNewRow) continue;

                    var cell = row.Cells[0];
                    if (cell?.Value != null && cell.Value.ToString().Contains(partNumber))
                    {
                        test2 = true;
                        partNumberFinal = partNumber;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in CheckIfPNExists (Columbus data): " + ex.Message);
            }

            try
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.IsNewRow) continue;

                    var mapCell = row.Cells[0];
                    var mappedNameCell = row.Cells[1];

                    if (mapCell?.Value != null && mappedNameCell?.Value != null &&
                        mapCell.Value.ToString().Contains(partNumber))
                    {
                        test2 = true;
                        partNumberFinal = mappedNameCell.Value.ToString().ToUpper();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in CheckIfPNExists (mapping file): " + ex.Message);
            }

            MessageBox.Show($"{partNumber} could not be found in Columbus database!");

            DialogResult dialogResult = MessageBox.Show("Would you like to import part into Columbus?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ImportToColumbus(partNumber);
            }

            partNumberFinal = null;
        }
        //This was a work in progress on allowing user to import a cad file into columbus and assigning it relevant part data
        private void ImportToColumbus(string columbusName)
        {
            MessageBox.Show("Please select drawing file.");
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.Filter = "AutoCad Drawings (*.DXF, *.DWG)|*.DXF;*.DWG";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string drawingName = openFileDialog1.FileName;
                string partName = columbusName;
                string pThickness = Microsoft.VisualBasic.Interaction.InputBox("Please enter part thickness:",
                        "User Input",
                        "",
                        200,
                        200);
                string pMaterial = Microsoft.VisualBasic.Interaction.InputBox("Please enter part material:",
                    "User Input",
                    "",
                    200,
                    200);
                dataGridView8.Rows[0].SetValues(partName, drawingName, pThickness, pMaterial);
                string csv = string.Empty;
                foreach (DataGridViewRow row in dataGridView8.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        csv += cell.Value.ToString().Replace(",", ";") + ',';
                    }
                    csv += "\r\n";
                }
                string folderPath = $"C:\\Users\\{Environment.UserName}\\Documents\\Columbus\\Imports\\";
                File.WriteAllText(folderPath + "New Part Import.csv", csv);
                dataGridView8.Rows.Clear();
                MessageBox.Show("Part imported into Columbus.  Re-input work order.");
                return;
            }
            else
            {
                MessageBox.Show("No drawing file selected");
                return;
            }
        }
        //Method to take columbus data and import it into table
        private void GetColumbusData(string partNumber, out string materialGrade, out string thickness)
        {
            DataGridViewRow row = dataGridView3.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Cells[0].Value.ToString().Equals(partNumber.ToUpper()))
                .First();
            int index = row.Index;
            materialGrade = dataGridView3.Rows[index].Cells[2].Value.ToString();
            thickness = dataGridView3.Rows[index].Cells[1].Value.ToString();
        }
        //takes info from import table and converts to a csv file to be imported into columbus
        private void SendToColumbus()
        {
            string csv = string.Empty;
            if (dataGridView1.Rows[0].Cells[0].Value == null || dataGridView1.Rows[0].Cells[0].Value.ToString() == "")
            {
                MessageBox.Show("No work orders selected for import!");
                return;
            }         
                   foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    csv += cell.Value.ToString().Replace(",", ";") + ',';
                }
                csv += "\r\n";
            }
            string folderPath = $"C:\\Users\\{Environment.UserName}\\Documents\\Columbus\\Imports\\";
            File.WriteAllText(folderPath + "WO Import.csv", csv);
            dataGridView1.Rows.Clear();
        }
                       
        public Form1()
        {
            InitializeComponent();
            dataGridView1.Rows.Add();
            textBoxDbLocation.Text = Properties.Settings.Default.nestingLocation;
            PopulateColumbusData();
            PopulateMasData();
            InsertDgvIntoForm();
            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.fileLocation) || !Directory.Exists(Properties.Settings.Default.fileLocation))
            {
                MessageBox.Show("Please select a location to store your part match file (User Settings.xml).");

                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select folder for part match file";

                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        Properties.Settings.Default.fileLocation = folderDialog.SelectedPath;
                        Properties.Settings.Default.Save();
                    }
                    else
                    {
                        MessageBox.Show("A valid location is required to run the program.");
                        Application.Exit();
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportDgvToXML();
            panel4.Visible = false;
            MessageBox.Show("Part list has been updated.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Please enter work order number!");
                return;
            }
            PopulateColumbusData();
            PopulateMasData();
            InputWOInfo();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SendToColumbus();
            dataGridView1.Rows.Add();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                Properties.Settings.Default.fileLocation = folderBrowserDialog1.SelectedPath;
                Properties.Settings.Default.Save();
            }
            InsertDgvIntoForm();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
            {
                button1_Click(this, new EventArgs());
            }
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.dataGridView1.Rows[e.RowIndex].Selected = true;
                this.rowIndex = e.RowIndex;
                this.dataGridView1.CurrentCell = this.dataGridView1.Rows[e.RowIndex].Cells[1];
                this.contextMenuStrip1.Show(this.dataGridView1, e.Location);
                contextMenuStrip1.Show(Cursor.Position);
            }
        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            if (!this.dataGridView1.Rows[this.rowIndex].IsNewRow)
            {
                this.dataGridView1.Rows.RemoveAt(this.rowIndex);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void buttonOpenSettings_Click(object sender, EventArgs e)
        {
            SettingsForm settingsForm = new SettingsForm();
            settingsForm.ShowDialog();
            textBoxDbLocation.Text = Properties.Settings.Default.nestingLocation;
        }
    }
}
