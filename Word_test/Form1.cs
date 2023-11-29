using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Word_test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        string str;
        private void button1_Click(object sender, EventArgs e)
        {


            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var helper = new WordHelper(dialog.FileName);
                    var items = new Dictionary<string, string>
                    {
                        {"<Organization>", textBox1.Text},
                        {"<FIO_dir>", textBox2.Text},
                        {"<Job_person>", textBox3.Text},
                        {"<FIO_person>", textBox4.Text},
                        {"<Date_start>", dateTimePicker1.Text},
                        {"<Date_end>", dateTimePicker2.Text},
                        {"<Value_date>", numericUpDown1.Text},
                        {"<Date_now>", dateTimePicker3.Text},
                    };

                    helper.Process(items);
                    str = items.GetString();
                    SQLka.SQLk(str);
                }
            }
            

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            


        }
        Regex _regToOrg = new Regex(@"Генеральному директору (?<Organization>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regToDir = new Regex(@"организации (?<FIO_dir>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regJob = new Regex(@"на должности (?<Job_person>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regPerson = new Regex(@"От (?<FIO_person>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regDateStart = new Regex(@"с (?<Date_start>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regDateEnd = new Regex(@" по (?<Date_end>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regValueDate = new Regex(@"сроком на (?<Value_date>.+)", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex _regDateNow = new Regex(@"(?<Date_now>[0-9]+\.[0-9]+\.[0-9]+)г", RegexOptions.Compiled | RegexOptions.Singleline);
        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var helper = new WordHelper(dialog.FileName);
                        List<string> textBlocks = helper.ReadText();
                        
                        string toorg=null, todir=null, job=null, person=null, datestrart=null, dateend = null, valuedate = null, datenow=null;

                        foreach (var block in textBlocks.Where(b => ! string.IsNullOrEmpty(b)))
                        {
                            Match m = null;
                            if (string.IsNullOrEmpty(toorg))
                            {
                                m = _regToOrg.Match(block);
                                if (m != null && m.Success)
                                {
                                    toorg = m.Groups["Organization"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(todir))
                            {
                                m = _regToDir.Match(block);
                                if (m != null && m.Success)
                                {
                                    todir = m.Groups["FIO_dir"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(job))
                            {
                                m = _regJob.Match(block);
                                if (m != null && m.Success)
                                {
                                    job = m.Groups["Job_person"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(person))
                            {
                                m = _regPerson.Match(block);
                                if (m != null && m.Success)
                                {
                                    person = m.Groups["FIO_person"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(datestrart))
                            {
                                m = _regDateStart.Match(block);
                                if (m != null && m.Success)
                                {
                                    datestrart = m.Groups["Date_start"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(dateend))
                            {
                                m = _regDateEnd.Match(block);
                                if (m != null && m.Success)
                                {
                                    dateend = m.Groups["Date_end"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(valuedate))
                            {
                                m = _regValueDate.Match(block);
                                if (m != null && m.Success)
                                {
                                    valuedate = m.Groups["Value_date"].Value;
                                }
                            }

                            if (string.IsNullOrEmpty(datenow))
                            {
                                m = _regDateNow.Match(block);
                                if (m != null && m.Success)
                                {
                                    datenow = m.Groups["Date_now"].Value;
                                }
                            }

                            Console.WriteLine($"toorg = {toorg}, todir = {todir}, job = {job}, person = {person}, datestart = {datestrart}, dateend = {dateend}, valuedate = {valuedate}, datenow = {datenow}");
                        }
                        
                    }
                }
            }
            catch { }
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
                EmailService emailService = new EmailService();
                emailService.SendEmail(textBox5.Text, "Тема письма",str);
                
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
