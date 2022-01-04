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
//change the dtStartYear to the new start day of the school year and make it a sunday

namespace Student_awards
{
    public partial class TestGoogle : Form
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string sheet = "Sheet1";
        static string Data = "Data";
        static SheetsService Service;

        public TestGoogle()
        {
            InitializeComponent();

            grid.ColumnCount = 5;
            grid.Columns[0].Name = "Names";
            grid.Columns[1].Name = "Star A";
            grid.Columns[2].Name = "Star B";
            grid.Columns[3].Name = "Star C";
            grid.Columns[4].Name = "Star D";
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            grid.Columns.Add(btn);
            btn.HeaderText = "A Stars";
            btn.Text = "A";
            btn.Name = "btn";
            btn.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn2 = new DataGridViewButtonColumn();
            grid.Columns.Add(btn2);
            btn2.HeaderText = "B Stars";
            btn2.Text = "B";
            btn2.Name = "btn2";
            btn2.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn3 = new DataGridViewButtonColumn();
            grid.Columns.Add(btn3);
            btn3.HeaderText = "C Stars";
            btn3.Text = "C";
            btn3.Name = "btn3";
            btn3.UseColumnTextForButtonValue = true;
            DataGridViewButtonColumn btn4 = new DataGridViewButtonColumn();
            grid.Columns.Add(btn4);
            btn4.HeaderText = "D Stars";
            btn4.Text = "D";
            btn4.Name = "btn4";
            btn4.UseColumnTextForButtonValue = true;

        }
        //This is the method for when 1 of the 4 buttons is clicked it shows the student then sends the info to the UpdateEntry 
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 5)
                {
                    MessageBox.Show(grid.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an A star");
                    UpdateEntry("B", e.RowIndex + 1, (Int32.Parse(grid.Rows[e.RowIndex].Cells[1].Value.ToString()) + 1).ToString());
                    UpdateData(grid.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an A star");
                }
                if (e.ColumnIndex == 6)
                {
                    MessageBox.Show(grid.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an B star");
                    UpdateEntry("C", e.RowIndex + 1, (Int32.Parse(grid.Rows[e.RowIndex].Cells[2].Value.ToString()) + 1).ToString());
                    UpdateData(grid.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an B star");
                }
                if (e.ColumnIndex == 7)
                {
                    MessageBox.Show(grid.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an C star");
                    UpdateEntry("D", e.RowIndex + 1, (Int32.Parse(grid.Rows[e.RowIndex].Cells[3].Value.ToString()) + 1).ToString());
                    UpdateData(grid.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an C star");
                }
                if (e.ColumnIndex == 8)
                {
                    MessageBox.Show(grid.Rows[e.RowIndex].Cells[0].Value.ToString() + " got awarded an D star");
                    UpdateEntry("E", e.RowIndex + 1, (Int32.Parse(grid.Rows[e.RowIndex].Cells[4].Value.ToString()) + 1).ToString());
                    UpdateData(grid.Rows[e.RowIndex].Cells[0].Value.ToString(), "Got awarded an D star");
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
                //MessageBox.Show("Changed to: " + comboBox1.SelectedItem.ToString());
                ReadEntries();
            }
            
        }
        //This is the start method and is responsable for connecting to the api and filling the combobox
        private void TestGoogle_Load(object sender, EventArgs e)
        {
            GoogleCredential credential;

            using (var stream = new FileStream("Key.json", FileMode.Open,FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName,});

            var ssRequest = Service.Spreadsheets.Get(SpreadsheetId);
            var ss = ssRequest.Execute();
            List<string> sheetList = new List<string>();

            foreach (var sheet in ss.Sheets)
            {
                sheetList.Add(sheet.Properties.Title);
                //MessageBox.Show(sheet.Properties.Title);
            }
            BindingSource bs = new BindingSource();
            sheetList.Remove("Data");
            sheetList.Remove("Teacher Data");
            bs.DataSource = sheetList;
            comboBox1.DataSource = bs;
            comboBox1.SelectedIndex=-1;
            grid.Rows.Clear();
            //ReadEntries();
        }

        void ReadEntries()
        {
            grid.Rows.Clear();
            var range = $"{sheet}!A1:E40";
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
                            string[] row = new string[] { Row[0].ToString(), Row[1].ToString(), Row[2].ToString(), Row[3].ToString(), Row[4].ToString() };
                            grid.Rows.Add(row);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The student "+ Row[0].ToString()+" does not have a value in the database");
                            
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

        void UpdateEntry(string coloum, int row, string value)
        {
            var range = $"{sheet}!"+coloum+row;
            var valueRange = new ValueRange();
            var objectList = new List<object>() { value};
            valueRange.Values = new List<IList<object>> { objectList};
            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            ReadEntries();

        }
        //change the dtStartYear to the new start day of the school year and make it a sunday
        void UpdateData(string Name, string Star)
        {
            DateTime dtToday = DateTime.Now;
            DateTime dtStartYear = new System.DateTime(2021, 8, 29, 0, 0, 0);
            TimeSpan diffResult = dtToday.Subtract(dtStartYear);


            var range = $"{Data}!A:G";
            var valueRange = new ValueRange();

            var objectList = new List<object>() {"Week"+ ((int)diffResult.TotalDays / 7 + 1), DateTime.Now.ToString("D"), Name,sheet,Star, DateTime.Now.ToString("t"), Login.Username };
            valueRange.Values = new List<IList<object>> { objectList };
            var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
            

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
