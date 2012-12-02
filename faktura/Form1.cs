using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.XPath;
using System.Diagnostics;



namespace faktura
{

    public partial class Form1 : Form
    {



        public string cestaKFakturam;
        public string cestaKXmlkam;
        public Obsah o;

        public Form1()
        {
            InitializeComponent();
            string homePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            cestaKFakturam = homePath + @"\Dropbox\fak\";
            cestaKXmlkam = homePath+ @"\Downloads\";

            
            zobraz();



        }

        private void zmaz()
        {
            foreach (object c in this.Controls)
            {
                if (c is TextBox)
                {
                    ((TextBox)c).Text = "";
                }
            }
        }

        private void zobraz()
        {


            try
            {

                XPathDocument doc = new XPathDocument(najdiXmlko(cestaKXmlkam));
                XPathNavigator navigator = doc.CreateNavigator();


                textBox2.Text = nacitajParameter(navigator, "/FILE/ZMLUVA/CENAO").Replace(".00", "");
                textBox3.Text = nacitajParameter(navigator, "/FILE/ZMLUVA/SPLODB");
                textBox4.Text = nacitajParameter(navigator, "/FILE/ZMLUVA/EVIDCO");

                textBox5.Text = nacitajParameter(navigator, "/FILE/ZMLADR/STAT", 1) + " " + nacitajParameter(navigator, "/FILE/ZMLADR/MESTO", 1);
                textBox6.Text = nacitajParameter(navigator, "/FILE/ZMLADR/STAT", 2) + " " + nacitajParameter(navigator, "/FILE/ZMLADR/MESTO", 2);

                textBox7.Text = nacitajParameter(navigator, "/FILE/ORGAN/ORGANOF") + "\r\n" +
                    nacitajParameter(navigator, "/FILE/ORGAN/ORGAN3") + "\r\n" +
                    nacitajParameter(navigator, "/FILE/ORGAN/ORGAN4") + "\r\n" +
                    nacitajParameter(navigator, "/FILE/ORGAN/PSC") + "\r\n" +
                    "IČO " + nacitajParameter(navigator, "/FILE/ORGAN/ICO") + "\r\n" +
                    "DIČ " + nacitajParameter(navigator, "/FILE/ORGAN/DIC") + "\r\n" +
                    "IČ DPH " + nacitajParameter(navigator, "/FILE/ORGAN/VAT") + "\r\n";

                DateTime d = DateTime.ParseExact(nacitajParameter(navigator, "/FILE/ZMLADR/TERMDAT", 1), "yyyy-MM-dd", null);
                textBox8.Text = d.ToString("dd.MM.yyyy");

                textBox1.Text = (1 + zistiCisloOstatnej(cestaKFakturam)).ToString("000");

                o = new Obsah();

                o.rok = DateTime.ParseExact(nacitajParameter(navigator, "/FILE/ZMLADR/TERMDAT", 1), "yyyy-MM-dd", null).ToString("yyyy");
                o.cislo_faktury_trojmiestne = (1 + zistiCisloOstatnej(cestaKFakturam)).ToString("000");
                o.cislo_objednavky = nacitajParameter(navigator, "/FILE/ZMLUVA/EVIDCO");
                o.splatnost_pocet_dni = nacitajParameter(navigator, "/FILE/ZMLUVA/SPLODB");
                o.datum_dodania_sluzby_dd_MM_yyyy = DateTime.ParseExact(nacitajParameter(navigator, "/FILE/ZMLADR/TERMDAT", 1), "yyyy-MM-dd", null).ToString("dd.MM.yyyy");

                o.adresa = "<text:p text:style-name=\"P1\">" + nacitajParameter(navigator, "/FILE/ORGAN/ORGANOF") + "</text:p>" + "\r\n" +
                    "<text:p text:style-name=\"P1\">" + nacitajParameter(navigator, "/FILE/ORGAN/ORGAN3") + "</text:p>" + "\r\n" +
                    "<text:p text:style-name=\"P1\">" + nacitajParameter(navigator, "/FILE/ORGAN/ORGAN4") + "</text:p>" + "\r\n" +
                    "<text:p text:style-name=\"P1\">" + nacitajParameter(navigator, "/FILE/ORGAN/PSC") + "</text:p>" + "\r\n" +
                    "<text:p text:style-name=\"P1\"/>\r\n" +
                    "<text:p text:style-name=\"P1\">" + "IČO " + nacitajParameter(navigator, "/FILE/ORGAN/ICO") + "</text:p>" + "\r\n" +
                    "<text:p text:style-name=\"P1\">" + "DIČ " + nacitajParameter(navigator, "/FILE/ORGAN/DIC") + "</text:p>" + "\r\n" +
                    "<text:p text:style-name=\"P1\">" + "IČ DPH " + nacitajParameter(navigator, "/FILE/ORGAN/VAT") + "</text:p>" + "\r\n";

                o.cena_bez_desatinnej_casti = nacitajParameter(navigator, "/FILE/ZMLUVA/CENAO").Replace(".00", "");
                o.miesto_nakladky = nacitajParameter(navigator, "/FILE/ZMLADR/STAT", 1) + " " + nacitajParameter(navigator, "/FILE/ZMLADR/MESTO", 1);
                o.miesto_vykladky = nacitajParameter(navigator, "/FILE/ZMLADR/STAT", 2) + " " + nacitajParameter(navigator, "/FILE/ZMLADR/MESTO", 2);








            }
            catch (Exception ee)
            {

                MessageBox.Show(ee.Message);
            }
        }

        private string nacitajParameter(XPathNavigator navigator, string p, int p_2)
        {
            string vystup = "";


            XPathNodeIterator iter = navigator.Select(p);
            if (p_2 == 1)
            {
                iter.MoveNext();
                vystup = iter.Current.Value;

            }
            if (p_2 == 2)
            {
                while (iter.MoveNext())
                {
                    vystup = iter.Current.Value;
                }
            }
            return vystup;
        }

        private static string nacitajParameter(XPathNavigator navigator, string cestavXML)
        {
            string vystup = "";


            XPathNodeIterator iter = navigator.Select(cestavXML);
            while (iter.MoveNext())
            {
                vystup = iter.Current.Value;
            }
            return vystup;
        }

        private string najdiXmlko(string cestaKXmlkam)
        {
            string cesta = "";
            DirectoryInfo zdrojAdr = new DirectoryInfo(cestaKXmlkam);

            if (zdrojAdr.Exists)
            {
                FileInfo[] odtcka = zdrojAdr.GetFiles("*.xml", SearchOption.TopDirectoryOnly);
                foreach (FileInfo f in odtcka)
                {
                    return f.FullName;
                    // tu bude este presun XML-ka
                }
            }
            else throw new Exception("neexistuje cesta:" + cestaKXmlkam);
            throw new Exception("neexistuje xml subor na ceste:" + cestaKXmlkam);

        }

        private int zistiCisloOstatnej(string cestaKFakturam)
        {
            int max = 0;
            DirectoryInfo zdrojAdr = new DirectoryInfo(cestaKFakturam);

            if (zdrojAdr.Exists)
            {
                FileInfo[] odtcka = zdrojAdr.GetFiles("*.odt");
                foreach (FileInfo f in odtcka)
                {
                    if (max < int.Parse(f.Name.Split(new char[] { '_' })[1]))
                        max = int.Parse(f.Name.Split(new char[] { '_' })[1]);
                }
            }

            return max;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            zobraz();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            zmaz();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string a = nacitaj_template(cestaKFakturam + "content_template.xml");

            a = a.Replace("--rok--", o.rok);
            a = a.Replace("--cislo_faktury_trojmiestne--", o.cislo_faktury_trojmiestne);
            a = a.Replace("--cislo_objednavky--", o.cislo_objednavky);
            a = a.Replace("--splatnost_pocet_dni--", o.splatnost_pocet_dni);
            a = a.Replace("--datum_dodania_sluzby_dd.MM.yyyy--", o.datum_dodania_sluzby_dd_MM_yyyy);
            a = a.Replace("--adresa--", o.adresa);
            a = a.Replace("--cena_bez_desatinnej_casti--", o.cena_bez_desatinnej_casti);
            a = a.Replace("--miesto_nakladky--", o.miesto_nakladky);
            a = a.Replace("--miesto_vykladky--", o.miesto_vykladky);

            StreamWriter sw = new StreamWriter(cestaKFakturam + "content.xml");
            sw.Write(a);
            sw.Close();

            File.Copy(cestaKFakturam+"FAKTURA_Template.zip",
                cestaKFakturam+"FAKTURA_"+o.cislo_faktury_trojmiestne+"_" +o.rok+".zip");//"Template.zip"

            ProcessStartInfo p = new ProcessStartInfo(cestaKFakturam + "zip.exe", " -m FAKTURA_"+o.cislo_faktury_trojmiestne+"_" +o.rok+".zip content.xml");
            p.WorkingDirectory = cestaKFakturam;
            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(p))
                {
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                // Log error.
            }

            File.Move(cestaKFakturam + "FAKTURA_" + o.cislo_faktury_trojmiestne + "_" + o.rok + ".zip",
                cestaKFakturam + "FAKTURA_" + o.cislo_faktury_trojmiestne + "_" + o.rok + ".odt");

            Process.Start(cestaKFakturam + "FAKTURA_" + o.cislo_faktury_trojmiestne + "_" + o.rok + ".odt");



        }

        private string nacitaj_template(string p)
        {
            string templ = "";
            StreamReader sr = new StreamReader(p);
            templ = sr.ReadToEnd();

            return templ;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DirectoryInfo d = new DirectoryInfo(cestaKXmlkam + "spracovane\\");
            if (!d.Exists)
                d.Create();
            File.Move(najdiXmlko(cestaKXmlkam), d.FullName + najdiXmlko(cestaKXmlkam).Split(new char[]{'\\'})[najdiXmlko(cestaKXmlkam).Split(new char[]{'\\'}).Length-1]);
 
        }
    }

    public class Obsah
    {
        public string rok;
        public string cislo_faktury_trojmiestne;
        public string cislo_objednavky;
        public string splatnost_pocet_dni;
        public string datum_dodania_sluzby_dd_MM_yyyy;
        public string adresa;
        public string cena_bez_desatinnej_casti;
        public string miesto_nakladky;
        public string miesto_vykladky;

    }
}
