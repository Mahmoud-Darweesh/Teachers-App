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
    public partial class Notifications : Form
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string sheet = "Sheet1";
        static string Data = "Data";
        static SheetsService Service;
        public Notifications()
        {
            InitializeComponent();
        }

        private void Notifications_Load(object sender, EventArgs e)
        {
            GoogleCredential credential;

            using (var stream = new FileStream(Path.GetFullPath("Key.json"), FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });

            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].Name = "Names";
            dataGridView1.Columns[1].Name = "Notification";
            dataGridView1.Columns[2].Name = "Subject HoD";
            dataGridView1.Columns[3].Name = "Notification ID";

            ReadEntries();
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
            
            this.BackColor = ThemeColor.SecondaryColor;
        }

        void ReadEntries()
        {
            dataGridView1.Rows.Clear();
            var range = $"{"Notifications"}!A2:D";
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
                            Row[3].ToString();
                        }
                        catch (Exception)
                        {

                            string[] row = new string[] { Row[0].ToString(), Row[1].ToString(), Row[2].ToString(), values.IndexOf(Row).ToString() };


                            dataGridView1.Rows.Add(row);
                        }
                        
                    }
                }
                else
                {
                    MessageBox.Show("No new notifications!");
                }

            }
            catch (Exception)
            {
            }

        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1 && e.RowIndex >=0)
            {
                saveFileDialog1.Title = "Save HPS Certificate";
                saveFileDialog1.FileName = dataGridView1[0, e.RowIndex].Value.ToString() + "'s " + dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() + " Certificate";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Certificate.CreateWordDocument(Certificate.doc, saveFileDialog1.FileName, dataGridView1[0, e.RowIndex].Value.ToString(), dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString(), dataGridView1[2, e.RowIndex].Value.ToString());
                    //MessageBox.Show(dataGridView1[3, e.RowIndex].Value.ToString());
                    updateStatus(Int32.Parse(dataGridView1[3, e.RowIndex].Value.ToString()));
                }

                
            }
            
        }

        void updateStatus(int row)
        {
            var range = $"{"Notifications"}!" + "D" + (row+2);
            var valueRange = new ValueRange();
            var objectList = new List<object>() { "Done" };
            valueRange.Values = new List<IList<object>> { objectList };
            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            ReadEntries();
        }
    }
    
}
