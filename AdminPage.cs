using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using DGVPrinterHelper;

namespace Student_awards
{
    public partial class AdminPage : Form
    {
        //Variables
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string Data = "Data";
        static bool isClass;
        static SheetsService Service;
        static string Pic = Path.GetFullPath("HR_LogoSml.png");
        static bool advanced = false;


        public AdminPage()
        {
            InitializeComponent();
        }

        //Start of form
        private void AdminPage_Load(object sender, EventArgs e)
        {
            //Start of the API
            GoogleCredential credential;

            using (var stream = new FileStream(Path.GetFullPath("Key.json"), FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });
            //Populating the main combobox
            BindingSource bs = new BindingSource();
            List<string> displayOptions = new List<string>();
            bs.DataSource = displayOptions;
            displayOptions.Add("Classes");
            displayOptions.Add("Students");
            comboBox1.DataSource = bs;
            //Clearing the first choice 
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            //Hiding the advanced section
            comboBox3.Hide();
            comboBox4.Hide();
            comboBox5.Hide();
            comboBox6.Hide();
            button4.Hide();
            button5.Hide();
            label2.Hide();
            label3.Hide();
            label4.Hide();
            label5.Hide();
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

        //When the user chooses an item in the main combobox it populats the second combobox
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            if (comboBox1.SelectedIndex == -1)
            {

            }
            else
            {
                if (comboBox1.SelectedItem.ToString() == "Classes")
                {
                    Classes();
                }
                else if (comboBox1.SelectedItem.ToString() == "Students")
                {
                    Students();
                }
            }
        }
        //If class was chosen in the main combobox then the second combobox is populated accordingly
        void Classes()
        {
            isClass = true;
            comboBox2.Items.Clear();
            foreach (var item in Form1.Grades)
            {
                comboBox2.Items.Add(item);
            }
        }
        //If students was chosen in the main combobox then the second combobox is populated accordingly
        void Students()
        {
            //Gets all the grades
            isClass = false;           
            comboBox2.Items.Clear();
            foreach (var item in Form1.Students)
            {
                comboBox2.Items.Add(item);
            }
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
        }
        //If the user chooses an item from the second combobox the grid is filled accordingly 
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //if they choose class in the main combobox then the datagrid will populate accoringly
            if (isClass)
            {
                dataGridView1.Rows.Clear();
                var range = $"{comboBox2.SelectedItem.ToString()}!A1:H";
                var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;
                //creates headers
                dataGridView1.ColumnCount = 8;
                dataGridView1.Columns[0].Name = "Names";
                dataGridView1.Columns[1].Name = "Star A";
                dataGridView1.Columns[2].Name = "Star B";
                dataGridView1.Columns[3].Name = "Star C";
                dataGridView1.Columns[4].Name = "Star D";
                dataGridView1.Columns[5].Name = "Star E";
                dataGridView1.Columns[6].Name = "Star F";
                dataGridView1.Columns[7].Name = "Score";
                try
                {
                    if (values != null && values.Count > 0)
                    {
                        foreach (var Row in values)
                        {
                            try
                            {
                                string[] row = new string[] { Row[0].ToString(), Row[1].ToString(), Row[2].ToString(), Row[3].ToString(), Row[4].ToString(), Row[5].ToString(), Row[6].ToString(), Row[7].ToString() };
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
            else if (!isClass)
            {
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();
                var range = $"{Data}!A1:H";
                var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;
                dataGridView1.ColumnCount = 7;
                dataGridView1.Columns[0].Name = "Week";
                dataGridView1.Columns[1].Name = "Date";
                dataGridView1.Columns[2].Name = "Name";
                dataGridView1.Columns[3].Name = "Grade";
                dataGridView1.Columns[4].Name = "Star";
                //dataGridView1.Columns[5].Name = "Time";
                dataGridView1.Columns[5].Name = "Teacher";
                dataGridView1.Columns[6].Name = "Score";
                try
                {
                    if (values != null && values.Count > 0)
                    {
                        foreach (var Row in values)
                        {
                            if (Row[2].ToString() == comboBox2.SelectedItem.ToString())
                            {
                                try
                                {
                                    string[] row = new string[] { Row[0].ToString(), Row[1].ToString(), Row[2].ToString(), Row[3].ToString(), Row[4].ToString(), Row[6].ToString(), Row[7].ToString() };
                                    dataGridView1.Rows.Add(row);
                                }
                                catch (Exception)
                                {
                                    MessageBox.Show("The student " + Row[2].ToString() + " does not have a value in the database");
                                }
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
        }
        //Prints the data
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                DGVPrinter.ImbeddedImage img1 = new DGVPrinter.ImbeddedImage();
                DGVPrinter printer = new DGVPrinter();
                Bitmap bitmap1 = new Bitmap(Pic);
                img1.theImage = bitmap1; img1.ImageX = 5; img1.ImageY = 10;
                img1.ImageAlignment = DGVPrinter.Alignment.Left;
                img1.ImageLocation = DGVPrinter.Location.Header;
                printer.ImbeddedImageList.Add(img1);
                if (button3.Text == "Advanced")
                {
                    printer.Title = comboBox2.SelectedItem.ToString() + "'s Report";
                }
                else
                {
                    printer.Title = "Custom Report";
                }
                printer.SubTitle = string.Format("Date: {0}", DateTime.Now.Date.ToString("D")) + "\n\n\n\n";
                printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
                printer.PageNumbers = true;
                printer.PageNumberInHeader = false;
                printer.PorportionalColumns = true;
                printer.HeaderCellAlignment = StringAlignment.Near;
                printer.Footer = "HPS Preformance Level Sheet";
                printer.FooterSpacing = 15;
                printer.PorportionalColumns = false;
                printer.PrintMargins.Left = 50;
                printer.PrintMargins.Right = 30;
                //foreach (DataGridViewColumn c in dataGridView1.Columns)
                //    if (c.Width > 140)
                //    {
                //        c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                //        c.Width = 140;
                //    }
                printer.PrintDataGridView(dataGridView1);
            }
            catch (Exception)
            {
                MessageBox.Show("Load data before printing.");
            }
        }
        //goes back too preformance tab
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new PreformanceTab();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }
        //Changes what is shown or hidden
        private void button3_Click(object sender, EventArgs e)
        {
            if (!advanced)
            {
                comboBox1.Hide();
                comboBox2.Hide();
                comboBox3.Show();
                comboBox4.Show();
                comboBox5.Show();
                comboBox6.Show();
                button4.Show();
                button5.Show();
                label2.Show();
                label3.Show();
                label4.Show();
                label5.Show();
                weekLoad();
                studentLoad();
                gradeLoad();
                teacherLoad();
                button3.Text = "Basic";
                advanced = true;
            }
            else if (advanced)
            {
                comboBox1.Show();
                comboBox2.Show();
                comboBox3.Hide();
                comboBox4.Hide();
                comboBox5.Hide();
                comboBox6.Hide();
                button4.Hide();
                button5.Hide();
                label2.Hide();
                label3.Hide();
                label4.Hide();
                label5.Hide();
                button3.Text = "Advanced";
                advanced = false;
            }

        }
        //goes through all the weeks and takes the name of 1 of the weeks so there are no duplicates
        void weekLoad()
        {
            comboBox3.Items.Clear();
            foreach (var item in Form1.Weeks)
            {
                comboBox3.Items.Add(item);
            }
        }

        void studentLoad()
        {
            comboBox4.Items.Clear();
            foreach (var item in Form1.Students)
            {
                comboBox4.Items.Add(item);
            }
        }

        void gradeLoad()
        {
            comboBox5.Items.Clear();
            foreach (var item in Form1.Grades)
            {
                comboBox5.Items.Add(item);
            }
        }


        void teacherLoad()
        {
            comboBox6.Items.Clear();
            foreach (var item in Form1.Teachers)
            {
                comboBox6.Items.Add(item);
            }
        }
        //The load button dun dun duuuun this button first loads all the headers for the grid
        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            var range = $"{Data}!A1:H";
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            dataGridView1.ColumnCount = 7;
            dataGridView1.Columns[0].Name = "Week";
            dataGridView1.Columns[1].Name = "Date";
            dataGridView1.Columns[2].Name = "Name";
            dataGridView1.Columns[3].Name = "Grade";
            dataGridView1.Columns[4].Name = "Star";
            //dataGridView1.Columns[5].Name = "Time";
            dataGridView1.Columns[5].Name = "Teacher";
            dataGridView1.Columns[6].Name = "Score";
            try
            {
                if (values != null && values.Count > 0)
                {
                    foreach (var Row in values)
                    {
                        try
                        {
                            //here we see is the user left the feild emety then it is couning every thing for that field 
                            //as true and the only way for a row to be printed is for all 4 feilds to be true for that row
                            bool c1 = false;
                            bool c2 = false;
                            bool c3 = false;
                            bool c4 = false;

                            if (comboBox3.SelectedIndex == -1)
                            {
                                c1 = true;
                            }
                            else if (comboBox3.SelectedItem.ToString() == Row[0].ToString())
                            {
                                c1 = true;
                            }
                            if (comboBox4.SelectedIndex == -1)
                            {
                                c2 = true;
                            }
                            else if (comboBox4.SelectedItem.ToString() == Row[2].ToString())
                            {
                                c2 = true;
                            }
                            if (comboBox5.SelectedIndex == -1)
                            {
                                c3 = true;
                            }
                            else if (comboBox5.SelectedItem.ToString() == Row[3].ToString())
                            {
                                c3 = true;
                            }
                            if (comboBox6.SelectedIndex == -1)
                            {
                                c4 = true;
                            }
                            else if (comboBox6.SelectedItem.ToString() == Row[6].ToString())
                            {
                                c4 = true;
                            }

                            if (c1 && c2 && c3 && c4)
                            {
                                string[] row = new string[] { Row[0].ToString(), Row[1].ToString(), Row[2].ToString(), Row[3].ToString(), Row[4].ToString(), Row[6].ToString(), Row[7].ToString() };
                                dataGridView1.Rows.Add(row);
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
        }

        //reloads all the data
        private void button5_Click(object sender, EventArgs e)
        {
            weekLoad();
            studentLoad();
            gradeLoad();
            teacherLoad();
            comboBox3.ResetText();
            comboBox4.ResetText();
            comboBox5.ResetText();
            comboBox6.ResetText();
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
        }
    }
}
