using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using System.IO;
using System.Diagnostics;

namespace LaporanHasil
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.ActiveControl = textBox1;
            textBox1.Focus();
            DateTime dt1 = DateTime.Now;
            DateTime dt2 = DateTime.Parse("2019-07-01");

            if (dt1.Date > dt2.Date)
            {
                MessageBox.Show("Waktu aktivasi habis, contact @andibits : alvi.joan@gmail.com");
                this.Close();
            }
            StreamReader sr = new StreamReader("D:\\LaporanHasil/dir.txt");
            label13.Text = sr.ReadToEnd();
            sr.Close();
                   
        }

        OpenFileDialog ofd = new OpenFileDialog();

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                label13.Text = ofd.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string server = label13.Text;
            string file = "D:\\LaporanHasil/ExcelFile.xls";
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];

            sheet.Cells[sheet.Cells.LastRowIndex + 1, 0] = new Cell(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sheet.Cells[sheet.Cells.LastRowIndex, 1] = new Cell(textBox1.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 2] = new Cell(textBox2.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 3] = new Cell(textBox3.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 4] = new Cell(textBox4.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 5] = new Cell(textBox5.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 6] = new Cell(textBox6.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 7] = new Cell(textBox7.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 8] = new Cell(textBox8.Text);
            sheet.Cells[sheet.Cells.LastRowIndex, 9] = new Cell(textBox9.Text);

            book.Save(file);
            book.Save(server);

            StreamWriter go = new StreamWriter("D:\\LaporanHasil/dir.txt");
            go.Write(label13.Text);
            go.Close();

            String conString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file + ";Extended Properties=Excel 8.0;";

            OleDbConnection con = new OleDbConnection(conString);
            String sql = "SELECT * FROM [0]";

            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, con);
            DataSet ds = new DataSet();
            adapter.Fill(ds);
            
            dataGridView1.DataSource = ds.Tables[0];
            con.Close();
        }
    }
}
