using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;

namespace Student_awards
{
    public partial class PreformanceTab : Form
    {
        //Variables
        
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string sheet = "Sheet1";
        static string Data = "Data";
        static SheetsService Service;

        public PreformanceTab()
        {
            InitializeComponent();
            //Loads the headers
            bunifuDataGridView1.ColumnCount = 9;
            bunifuDataGridView1.Columns[0].Name = "Names";
            bunifuDataGridView1.Columns[1].Name = "Star A";
            bunifuDataGridView1.Columns[2].Name = "Star B";
            bunifuDataGridView1.Columns[3].Name = "Star C";
            bunifuDataGridView1.Columns[4].Name = "Star D";
            bunifuDataGridView1.Columns[5].Name = "Star E";
            bunifuDataGridView1.Columns[6].Name = "Star F";
            bunifuDataGridView1.Columns[7].Name = "Score";
            bunifuDataGridView1.Columns[8].Name = "Extra";
            foreach (DataGridViewColumn column in bunifuDataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //creates the buttons
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            bunifuDataGridView1.Columns.Add(btn);
            btn.HeaderText = "A Stars";
            btn.Text = "A";
            btn.Name = "btn";
            btn.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn2 = new DataGridViewButtonColumn();
            bunifuDataGridView1.Columns.Add(btn2);
            btn2.HeaderText = "B Stars";
            btn2.Text = "B";
            btn2.Name = "btn2";
            btn2.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn3 = new DataGridViewButtonColumn();
            bunifuDataGridView1.Columns.Add(btn3);
            btn3.HeaderText = "C Stars";
            btn3.Text = "C";
            btn3.Name = "btn3";
            btn3.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn4 = new DataGridViewButtonColumn();
            bunifuDataGridView1.Columns.Add(btn4);
            btn4.HeaderText = "D Stars";
            btn4.Text = "D";
            btn4.Name = "btn4";
            btn4.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn5 = new DataGridViewButtonColumn();
            bunifuDataGridView1.Columns.Add(btn5);
            btn5.HeaderText = "E Stars";
            btn5.Text = "E";
            btn5.Name = "btn5";
            btn5.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn6 = new DataGridViewButtonColumn();
            bunifuDataGridView1.Columns.Add(btn6);
            btn6.HeaderText = "F Stars";
            btn6.Text = "F";
            btn6.Name = "btn6";
            btn6.UseColumnTextForButtonValue = true;

        }
        //When a button is pressed it figures out which button from the columns and rows
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //first it fugures out the button then it displays the award given then it adds 1 to that value in the data base then it adds the action to the data file
                if (e.ColumnIndex == 9)
                {
                    MessageBox.Show(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an A star");
                    UpdateEntry("B", e.RowIndex + 1, (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()) + 1).ToString());
                    UpdateData(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an A star", bunifuDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString());
                    if (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()) >= 30)
                    {
                        addNotification(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(),"30A");
                    }
                    else if (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()) >= 20)
                    {
                        addNotification(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "20A");
                    }
                }
                if (e.ColumnIndex == 10)
                {
                    MessageBox.Show(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an B star");
                    UpdateEntry("C", e.RowIndex + 1, (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString()) + 1).ToString());
                    UpdateData(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an B star", bunifuDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString());
                    if (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString()) >= 20)
                    {
                        addNotification(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "20B");
                    }
                }
                if (e.ColumnIndex == 11)
                {
                    MessageBox.Show(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an C star");
                    UpdateEntry("D", e.RowIndex + 1, (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString()) + 1).ToString());
                    UpdateData(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an C star", bunifuDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString());
                    if (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString()) >= 10)
                    {
                        addNotification(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "10C");
                    }
                }
                if (e.ColumnIndex == 12)
                {
                    MessageBox.Show(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an D star");
                    UpdateEntry("E", e.RowIndex + 1, (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString()) + 1).ToString());
                    UpdateData(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an D star", bunifuDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString());
                }
                if (e.ColumnIndex == 13)
                {
                    MessageBox.Show(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an E star");
                    UpdateEntry("F", e.RowIndex + 1, (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString()) + 1).ToString());
                    UpdateData(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an E star", bunifuDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString());
                    if (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString()) >= 10)
                    {
                        addNotification(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "10E");
                    }
                }
                if (e.ColumnIndex == 14)
                {
                    MessageBox.Show(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an F star");
                    UpdateEntry("G", e.RowIndex + 1, (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString()) + 1).ToString());
                    UpdateData(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an F star", bunifuDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString());
                    if (Int32.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString()) >= 10)
                    {
                        addNotification(bunifuDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(), "10F");
                    }
                }
            }
            catch (Exception)
            {
            }

        }
        //This updates the datagrid when the combo box is updated
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
            }
            else
            {
                sheet = comboBox1.SelectedItem.ToString();
                bunifuDataGridView1.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.SortMode = DataGridViewColumnSortMode.NotSortable);
                ReadEntries();
            }

        }
        
        
        //populates the grid
        void ReadEntries()
        {
            bunifuDataGridView1.Rows.Clear();
            var range = $"{sheet}!A1:I";
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            try
            {
                if (values != null && values.Count > 0)
                {
                    foreach (var Row in values)
                    {
                        try
                        {
                            string[] row = new string[] { Row[0].ToString(), Row[1].ToString(), Row[2].ToString(), Row[3].ToString(), Row[4].ToString(), Row[5].ToString(), Row[6].ToString(), Row[7].ToString(), Row[8].ToString() };
                            bunifuDataGridView1.Rows.Add(row);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The student " + Row[0].ToString() + " does not have a value in the database");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Error: No data found");
                }
            }
            catch (Exception)
            {
            }




        }
        //adds 1 to the value or skill 
        void UpdateEntry(string coloum, int row, string value)
        {
            var range = $"{sheet}!" + coloum + row;
            var valueRange = new ValueRange();
            var objectList = new List<object>() { value };
            valueRange.Values = new List<IList<object>> { objectList };
            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            ReadEntries();
        }

        //change the dtStartYear to the new start day of the school year and make it a sunday
        void UpdateData(string Name, string Star, string score)
        {
            DateTime dtToday = DateTime.Now;
            DateTime dtStartYear = new System.DateTime(2021, 8, 29, 0, 0, 0);
            TimeSpan diffResult = dtToday.Subtract(dtStartYear);


            var range = $"{Data}!A:H";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Week " + ((int)diffResult.TotalDays / 7 + 1), DateTime.Now.ToString("D"), Name, sheet, Star, DateTime.Now.ToString("t"), Login.Username, score };
            valueRange.Values = new List<IList<object>> { objectList };
            var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }
        //This is the start method and is responsable for connecting to the api and filling the combobox
        private void PreformanceTab_Load(object sender, EventArgs e)
        {
            GoogleCredential credential;

            using (var stream = new FileStream(Path.GetFullPath("Key.json"), FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });

            foreach (var item in Form1.Grades)
            {
                comboBox1.Items.Add(item);
            }
            comboBox1.SelectedIndex = -1;
            bunifuDataGridView1.Rows.Clear();
            if (Login.Username == "Samer")
            {
            }
            else
            {
                //button1.Hide();
            }
            //label1.Text = "Hello, " + Login.Username;
            LoadTheme();
        }

        private void LoadTheme()
        {
            foreach (Control btns in this.Controls)
            {
                if (btns.GetType() == typeof(Button))
                {
                    Button btn = (Button)btns;
                    btn.BackColor = ThemeColor.PrimaryColor;
                    btn.ForeColor = System.Drawing.Color.White;
                    btn.FlatAppearance.BorderColor = ThemeColor.SecondaryColor;
                }
            }
            //bunifuDataGridView1.BackgroundColor = ThemeColor.SecondaryColor;
            this.BackColor = ThemeColor.SecondaryColor;
        }

        //switches to the admin page
        private void button1_Click(object sender, EventArgs e)
        {
            
            this.Hide();
            var form2 = new AdminPage();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }
        //goes to awards
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new Awards();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        void addNotification(string Name, string Notification)
        {
            String[] Titles = new String[8];

            var range = $"{"Options"}!A1:H";
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            try
            {
                if (values != null && values.Count > 0)
                {
                    foreach (var Row in values)
                    {
                        try
                        {
                            if (Row[0].ToString() == comboBox1.SelectedItem.ToString().Split(' ').Last())
                            {
                                for (int i = 1; i < 8; i++)
                                {
                                    Titles[i] = Row[i].ToString();
                                }
                            }

                        }
                        catch (Exception)
                        {


                        }

                    }
                }
            }
            catch (Exception)
            {


            }
            if (Titles[1] == null)
            {
                MessageBox.Show("Inform the IT that the name of the grade does not end with the subject of the options sheet has the subject name spelled incorrectly or their are no titles writen for the awards");
            }

            range = $"{"Notifications"}!A:C";
            var valueRange = new ValueRange();
            int titleId=1;
            switch (Notification)
            {
                case "20A":
                    titleId = 1;
                    break;
                case "30A":
                    titleId = 2;
                    break;
                case "20B":
                    titleId = 3;
                    break;
                case "10C":
                    titleId = 4;
                    break;
                case "10E":
                    titleId = 5;
                    break;
                case "10F":
                    titleId = 6;
                    break;


            }
            if (!notificationContains(Name, Titles[titleId]))
            {
                var objectList = new List<object>() { Name, Titles[titleId], Titles[7] };
                valueRange.Values = new List<IList<object>> { objectList };
                var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();
            }
            
        }

        static bool notificationContains(string name, string reason)
        {
            bool found = false;
            var range = $"{"Notifications"}!A1:D";
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            try
            {
                if (values != null && values.Count > 0)
                {
                    foreach (var Row in values)
                    {
                        try
                        {
                            if (name == Row[0].ToString())
                            {
                                if (reason == Row[1].ToString())
                                {
                                    found = true;
                                    return found;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The week " + Row[0].ToString() + " does not have a value in the database");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Error: No data found");
                }
            }
            catch (Exception)
            {
            }

            
                return false;
            
        }
    }
}
