using System;
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
using System.IO;

namespace Student_awards
{
    public partial class Login : Form
    {
        public static String[] Subjects = { "Linguistics", "Science", "Social studys", "Math", "Computer", "PE", "ADMIN" };
        public static string Username;
        public static string Subject;
        public static int SubjectId;
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Student_awards";
        static string SpreadsheetId = "1ixkcVNm-t8ZpDIdqeDdSSj2d6ZvbrhgInT1aBgTbsdw";
        static string TeacherData = "Teacher Data";
        //static string Data = "Data";
        static SheetsService Service;
        
        public Login()
        {
            InitializeComponent();
        }
        private void Login_Load(object sender, EventArgs e)
        {
            label1.Parent = pictureBox1;
            label1.BackColor = System.Drawing.Color.Transparent;
            
            GoogleCredential credential;
            System.Threading.Thread.Sleep(250);
            using (var stream = new FileStream("Key.json", FileMode.Open, FileAccess.Read))
            {
                System.Threading.Thread.Sleep(250);
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }
            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var range = $"{TeacherData}!A1:C";
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            bool found = false;
            try
            {
                if (values != null && values.Count > 0)
                {
                    foreach (var Row in values)
                    {
                        Console.WriteLine(Row[0].ToString());
                        Console.WriteLine(Row[1].ToString());
                        if (Row[0].ToString() == textBox1.Text)
                        {
                            if (Row[1].ToString() == textBox2.Text)
                            {
                                Username = Row[0].ToString();
                                Subject = Row[2].ToString();
                                SubjectId = Array.IndexOf(Subjects,Subject);
                                found = true;
                                this.Hide();
                                var form2 = new Dashboard();
                                form2.Closed += (s, args) => this.Close();
                                form2.Show();
                            }
                            else
                            {
                                MessageBox.Show("Password Incorrect!");

                            }
                        }



                    }
                }
                if (!found)
                {
                    MessageBox.Show("Username not found!");
                }
                




            }
            catch (Exception)
            {


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
