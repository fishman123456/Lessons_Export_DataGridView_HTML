using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Text.RegularExpressions;

namespace Export_DataGridView_HTML_CS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[3] { new DataColumn("Id", typeof(int)),
                                   new DataColumn("Name", typeof(string)),
                                   new DataColumn("Country",typeof(string)) });
            dt.Rows.Add(1, "John Hammond", "United States");
            dt.Rows.Add(2, "Mudassar Khan", "India");
            dt.Rows.Add(3, "Suzanne Mathews", "France");
            dt.Rows.Add(4, "Robert Schidner", "Russia");
            this.dataGridView1.DataSource = dt;
            this.dataGridView1.AllowUserToAddRows = false;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //Table start.
            string html = "<table cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-size: 9pt;font-family:arial'>";

            //Adding HeaderRow.
            html += "<tr>";
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                html += "<th style='background-color: #B8DBFD;border: 1px solid #ccc'>" + column.HeaderText + "</th>";
            }
            html += "</tr>";

            //Adding DataRow.
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                html += "<tr>";
                foreach (DataGridViewCell cell in row.Cells)
                {
                    html += "<td style='width:120px;border: 1px solid #ccc'>" + cell.Value.ToString() + "</td>";
                }
                html += "</tr>";
            }

            //Table end.
            html += "</table>";

            File.WriteAllText(@"first.html", html);
        }      
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void первая_Click(object sender, EventArgs e)
        {
            string[] dict = File.ReadAllLines("dict.txt", Encoding.GetEncoding(1251)); // Открываем словарь
            var text = File.ReadAllLines("text.txt", Encoding.GetEncoding(1251)); // Открываем текст
            const string boldItalic = "<b><i>{0}</i></b>";

            // Создаем и пишем в файл
            using (var writer = new StreamWriter("output.html", false, Encoding.GetEncoding(1251)))
            {
                writer.Write("<!DOCTYPE HTML><html><head><title>Text</title></head><body>");
                for (int index = 0; index < text.Length; index++)
                {
                    string str = text[index];
                    foreach (string word in dict)
                    {
                        string pattern = String.Format(@"\b{0}\b", word);
                        Regex regex = new Regex(pattern, RegexOptions.Compiled);
                        str = regex.Replace(str, match => String.Format(boldItalic, match.Value));
                    }

                    writer.Write(str);
                }


                writer.WriteLine("</body></html>");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
      
        }
    }
}
