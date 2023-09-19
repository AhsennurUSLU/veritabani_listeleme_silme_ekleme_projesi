using Azure;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Syncfusion.Pdf.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Xml.Linq;
using DataTable = System.Data.DataTable;
using System.Windows.Forms.VisualStyles;
using Document = iTextSharp.text.Document;
//using iTextSharp.text.pdf;
//using iTextSharp.text;

namespace excelProje
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // veritabanı bağlantısı
        SqlConnection baglan = new SqlConnection("Data Source=veritabanında yer alan bikgisayar adı;Initial Catalog=veri tabanı ismi;Integrated Security=True");
        
        private void verileriGoruntule()
        {
            listView1.Items.Clear();
            baglan.Open();
            SqlCommand komut = new SqlCommand("Select * From kitaplar", baglan);  // sql sorgusu ile veri çek
            SqlDataReader oku = komut.ExecuteReader();

            while (oku.Read())
            {
                ListViewItem ekle = new ListViewItem();
                ekle.Text = oku["id"].ToString();
                ekle.SubItems.Add(oku["kitapad"].ToString());
                ekle.SubItems.Add(oku["yazar"].ToString());
                ekle.SubItems.Add(oku["yayinevi"].ToString());
                ekle.SubItems.Add(oku["sayfa"].ToString());

                listView1.Items.Add(ekle);
            }
            baglan.Close();

        }

        // Verileri listelediğimiz buton kodları

        private void button1_Click(object sender, EventArgs e)
        {
            verileriGoruntule();
        }

        // verilere ekleme yaptığımız buton kodları (kaydet)
        private void button2_Click(object sender, EventArgs e)
        {
            baglan.Open();
            SqlCommand sorgu = new SqlCommand("Insert Into kitaplar (id,kitapad,yazar,yayinevi,sayfa) Values ('" + textBox1.Text.ToString() + "' , '" + textBox2.Text.ToString() + "' , '" + textBox3.Text.ToString() + "' , '" + textBox4.Text.ToString() + "' , '" + textBox5.Text.ToString() + "' )", baglan);
            sorgu.ExecuteNonQuery();
            baglan.Close();
            verileriGoruntule();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
        }
        int id = 0;

         
        // verileri sildiğimiz butona ait kodlar 
        private void button3_Click(object sender, EventArgs e)
        {
            baglan.Open();
            SqlCommand sorgu2 = new SqlCommand("Delete From kitaplar Where id =(" + id + ")", baglan);
            sorgu2.ExecuteNonQuery();
            baglan.Close();
            verileriGoruntule();

        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            id = int.Parse(listView1.SelectedItems[0].SubItems[0].Text);

            textBox1.Text = listView1.SelectedItems[0].SubItems[0].Text;
            textBox2.Text = listView1.SelectedItems[0].SubItems[1].Text;
            textBox3.Text = listView1.SelectedItems[0].SubItems[2].Text;
            textBox4.Text = listView1.SelectedItems[0].SubItems[3].Text;
            textBox5.Text = listView1.SelectedItems[0].SubItems[4].Text;
        }


        // EXCELE VERİ AKTARMA

        private void excelAktar(ListView lw, ProgressBar pb = null)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = xls.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)xls.ActiveSheet;  // çalışma alanın aktif olan çalışma alanı 

                xls.Visible = true;  // görünürlük aktif


                // alanları manuel olarak yazıyoruz
                /*
                ws.Cells[1, 1] = "Sıra";
                ws.Cells[1, 2] = "Kitap";
                ws.Cells[1, 3] = "Yazar";
                ws.Cells[1, 4] = "Yayınevi";
                ws.Cells[1, 5] = "Sayfa";
                */
                // eğer progress bar nesnesi null değil ise sıfırlama ve ayarlama işlemi gerçekleştir 
                if (pb != null)
                {
                    pb.Maximum = Convert.ToInt32(lw.Items.Count.ToString());
                    pb.Value = 0;

                }

                // şimdi dinamik olarak sütun bilgilerini alıyoruz.

                for (int x = 0; x < lw.Columns.Count; x++)
                {
                    // alanları manuel olarak yazıyoruz.
                    ws.Cells[1, x + 1] = lw.Columns[x].Text.ToString();

                }
                //  şimdi lw içerisindeki verileri dinamik  olarak aktarıyoruz.
                int i = 2; //2. satırdan itibaren içerikleri doldurmaya başla
                int j = 1;

                foreach (ListViewItem item in lw.Items)
                {
                    ws.Cells[i, j] = item.Text.ToString();
                    foreach (ListViewItem.ListViewSubItem subitem in item.SubItems)
                    {
                        ws.Cells[i, j] = subitem.Text.ToString();
                        j++;

                    }
                    j = 1;
                    i++;
                    // Eğer Progress bar nesnesi null değil ise artırma işlemini yap
                    if (pb != null)
                    {
                        pb.Value = i - 2;
                    }
                }
                // column sütunları yazı boyutuna göre ayarlıyor.
                xls.Columns.AutoFit();

                // aktarma işlemi sırasında alabileceğimiz hatalara karşı önlem olarak hata bastırma iişlemi yapılıyor.
                xls.AlertBeforeOverwriting = false;

            }
            catch (Exception)
            {
                // hata fırlatıyoruz
                throw;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            excelAktar(listView1);
        }





        // WORDE AKTARMA


        private void button5_Click(object sender, EventArgs e)
        {
            

            Microsoft.Office.Interop.Word.Application wrd = new Microsoft.Office.Interop.Word.Application();
            wrd.Visible = true;

            Microsoft.Office.Interop.Word.Document wrddoc;
            object wrdobj = System.Reflection.Missing.Value;
            wrddoc = wrd.Documents.Add(ref wrdobj);
            wrd.Selection.TypeText(listView1.Text);

            for (int i = 0; i < listView1.Items.Count; i++)
            {

                for (int j = 0; j < listView1.Items[i].SubItems.Count; j++)
                {
                    wrd.Selection.TypeText(listView1.Items[i].SubItems[j].Text + " " + "\n"); 


                }
            }


            wrd = null;
            
            
            

        }
       
      

    }
}
