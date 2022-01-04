using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
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

namespace Student_awards
{
    public partial class Form1 : Form
    {
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static SheetsService Service;
        public static List<string> Students = new List<string>();
        public static List<string> Weeks = new List<string>();
        public static List<string> Teachers = new List<string>();
        public static List<string> Grades = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            GoogleCredential credential;

            using (var stream = new FileStream(Path.GetFullPath("Key.json"), FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });


            studentLoad();
            weekLoad();
            teacherLoad();
            gradeLoad();
            bar();
            
        }

        void bar()
        {
            while (!(panel2.Width >= 700))
            {
                panel2.Width += 3;
            }
        }

        void studentLoad()
        {
            Students.Clear();
            var range = $"{"Data"}!A1:H";
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
                            if (!Students.Contains(Row[2].ToString()))
                            {
                                Students.Add(Row[2].ToString());
                                //MessageBox.Show(Row[2].ToString());
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The student " + Row[2].ToString() + " does not have a value in the database");
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

        void weekLoad()
        {
            Weeks.Clear();
            var range = $"{"Data"}!A1:H";
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
                            if (!Weeks.Contains(Row[0].ToString()))
                            {
                                Weeks.Add(Row[0].ToString());
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

        void teacherLoad()
        {
            Teachers.Clear();
            var range = $"{"Data"}!A1:H";
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
                            if (!Teachers.Contains(Row[6].ToString()))
                            {
                                Teachers.Add(Row[6].ToString());
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The Teacher " + Row[6].ToString() + " does not have a value in the database");
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

        void gradeLoad()
        {
            Grades.Clear();
            var ssRequest = Service.Spreadsheets.Get(SpreadsheetId);
            var ss = ssRequest.Execute();
            
            foreach (var sheet in ss.Sheets)
            {
                Grades.Add(sheet.Properties.Title);
            }
            
            Grades.Remove("Data");
            Grades.Remove("Teacher Data");
            Grades.Remove("Options");
            Grades.Remove("Notifications");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
            if (panel2.Width >= 700)
            {
                timer1.Stop();
                Form form2 = new Login();
                this.Hide();
                form2.Closed += (s, args) => this.Close();
                form2.Show();
            }
        }
    }
}
