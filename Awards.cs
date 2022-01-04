using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using DGVPrinterHelper;
using System.Linq;

namespace Student_awards
{
    public partial class Awards : Form
    {
        //Variables
        static String[] Titles = new String[8];
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string Pic = Path.GetFullPath("HR_LogoSml.png");
        static SheetsService Service;

        public Awards()
        {
            InitializeComponent();
        }
        //Begining of form
        private void Awards_Load(object sender, EventArgs e)
        {
            //Entering API
            GoogleCredential credential;

            using (var stream = new FileStream(Path.GetFullPath("Key.json"), FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });

            //Populating combobox
            classComboBox.SelectedIndex = -1;
            Classes();

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
        //Methods

        //Switch to the Preformance tab
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new PreformanceTab();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        //Prints the data
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                DGVPrinter.ImbeddedImage img1 = new DGVPrinter.ImbeddedImage();
                DGVPrinter printer = new DGVPrinter();
                Bitmap bitmap1 = new Bitmap(Pic);
                img1.theImage = bitmap1; img1.ImageX = 0; img1.ImageY = 10;
                img1.ImageAlignment = DGVPrinter.Alignment.Left;
                img1.ImageLocation = DGVPrinter.Location.Header;
                printer.ImbeddedImageList.Add(img1);
                printer.Title = classComboBox.SelectedItem.ToString() + "'s Rank Report";
                printer.SubTitle = string.Format("Date: {0}", DateTime.Now.Date.ToString("D")) + "\n\n\n\n";
                printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
                printer.PageNumbers = true;
                printer.PageNumberInHeader = false;
                printer.PorportionalColumns = true;
                printer.HeaderCellAlignment = StringAlignment.Near;
                printer.Footer = "HPS Preformance Level Sheet";
                printer.FooterSpacing = 15;
                printer.PrintDataGridView(dataGridView1);
            }
            catch (Exception)
            {
                MessageBox.Show("Load data before printing.");
            }

        }

        //Populating combobox
        void Classes()
        {
            classComboBox.Items.Clear();
            foreach (var item in Form1.Grades)
            {
                classComboBox.Items.Add(item);
            }

        }

        void rankss()
        {
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
                            if (Row[0].ToString() == classComboBox.SelectedItem.ToString().Split(' ').Last())
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
        }

        //When the item is changed in the combobox this method runs and updates the datagridview
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            rankss();
            //Setting up the gridview and requesting the range for the data
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Rows.Clear();
            var range = $"{classComboBox.SelectedItem.ToString()}!A1:H";
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            dataGridView1.ColumnCount = 7;
            dataGridView1.Columns[0].Name = "Names";
            dataGridView1.Columns[1].Name = "20A";
            dataGridView1.Columns[2].Name = "30A";
            dataGridView1.Columns[3].Name = "20B";
            dataGridView1.Columns[4].Name = "10C";
            dataGridView1.Columns[5].Name = "10E";
            dataGridView1.Columns[6].Name = "10F";
            String[] ranks = new String[7];
            try
            {
                if (values != null && values.Count > 0)
                {
                    foreach (var Row in values)
                    {
                        try
                        {
                            String[] row = new string[7];
                            row[0] = Row[0].ToString();
                            if (Int32.Parse(Row[1].ToString()) >= 20)
                            {
                                row[1] = Titles[1];
                            }
                            else
                            {
                                row[1] = "None";
                            }
                            if (Int32.Parse(Row[1].ToString()) >= 30)
                            {
                                row[2] = Titles[2];
                            }
                            else
                            {
                                row[2] = "None";
                            }
                            if (Int32.Parse(Row[2].ToString()) >= 20)
                            {
                                row[3] = Titles[3];
                            }
                            else
                            {
                                row[3] = "None";
                            }
                            if (Int32.Parse(Row[3].ToString()) >= 10)
                            {
                                row[4] = Titles[4];
                            }
                            else
                            {
                                row[4] = "None";
                            }
                            if (Int32.Parse(Row[5].ToString()) >= 10)
                            {
                                row[5] = Titles[5];
                            }
                            else
                            {
                                row[5] = "None";
                            }
                            if (Int32.Parse(Row[6].ToString()) >= 10)
                            {
                                row[6] = Titles[6];
                            }
                            else
                            {
                                row[6] = "None";
                            }
                                
                            //When all the ranks for the student are done they are printed on the grid then we loop back to the next student ^^ to the for loop
                            
                            dataGridView1.Rows.Add(row);
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex!=0)
            {
                saveFileDialog1.Title = "Save HPS Certificate";
                saveFileDialog1.FileName = dataGridView1[0, e.RowIndex].Value.ToString() + "'s " + dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() + " Certificate";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Certificate.CreateWordDocument(Certificate.doc, saveFileDialog1.FileName, dataGridView1[0, e.RowIndex].Value.ToString(), dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString(), Titles[7]);
                }

                //MessageBox.Show(dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString());
            }
            
        }
    }
}
