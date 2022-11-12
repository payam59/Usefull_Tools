using OfficeOpenXml;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace ConvertLatLongDec2UTM
{
    public partial class Form1 : Form
    {
        string coordinates=null;
        public Form1()
        {
            InitializeComponent();
        }

        /*
        #Simbolo 3932552000N 0024357950E 20
        Linea 334020999N 0565330008E 325820006N 0591200007E
        Linea 345423002N 0522023001E 333130998N 0535239999E
        Linea 325330000N 0545850000E 324008000N 0552339000E
        Simbolo 334020999N 0565330008E Tabas 2000#5
        Simbolo 325820006N 0591200007E birjand 2000#6


        */
        private string deglon(string deglon)
        {
           decimal signlon;
           if (deglon.Substring(0,1) == "-") { signlon = -1; } 
            else
            { signlon = 1; }

           decimal lonAbs = Math.Abs(Math.Round(Convert.ToDecimal(deglon) * 1000000));

           string converted = (Math.Floor(lonAbs / 1000000) * signlon).ToString() + "° " + Math.Floor(((lonAbs / 1000000) - Math.Floor(lonAbs / 1000000)) * 60).ToString() + "\' " + (Math.Floor(((((lonAbs / 1000000) - Math.Floor(lonAbs / 1000000)) * 60) - Math.Floor(((lonAbs / 1000000) - Math.Floor(lonAbs / 1000000)) * 60)) * 100000) * 60 / 100000).ToString() + "\"";
            return converted;
        
        }
        private string deglat(string deglat)
        {
            decimal signlat;
            if (deglat.Substring(0, 1) == "-") { signlat = -1; }
            else
            { signlat = 1; }

            decimal latAbs = Math.Abs(Math.Round(Convert.ToDecimal(deglat) * 1000000));

            string converted = ((Math.Floor(latAbs / 1000000) * signlat) + "° " + Math.Floor(((latAbs / 1000000) - Math.Floor(latAbs / 1000000)) * 60) + "\' " + (Math.Floor(((((latAbs / 1000000) - Math.Floor(latAbs / 1000000))* 60) - Math.Floor(((latAbs / 1000000) - Math.Floor(latAbs / 1000000)) * 60)) * 100000) * 60 / 100000) + '\"');
            return converted;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            string exfile;
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File";
            theDialog.Filter = "xlsx files|*.xlsx";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                exfile = theDialog.FileName;
                theDialog.InitialDirectory = Path.GetDirectoryName(theDialog.FileName); 
                using (var package = new ExcelPackage(exfile))
                {

                    var firstSheet = package.Workbook.Worksheets[0];
                    int colCount = firstSheet.Dimension.End.Column;  //get Column Count
                    int rowCount = firstSheet.Dimension.End.Row;

                    //                                          We will skip 1st row and only parse 2nd column with coordinates
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // for (int col = 1; col <= colCount; col++)
                        // {
                        //  302216N0481342E
                        string sampleout = "Simbolo 334020999N 0565330008E Tabas 2000#5";
                        string airportCode = firstSheet.Cells[row, 1].Value?.ToString().Trim();
                        string coordinate = firstSheet.Cells[row, 2].Value?.ToString().Trim();

                        string coordinateN = coordinate.Substring(0, coordinate.IndexOf('N'));
                        string coordinateE = coordinate.Substring(coordinate.IndexOf('N')+1).Replace("E","");

                        coordinates = coordinates + airportCode + "," + coordinateN + "," + coordinateE + ";";


                        string convE =deglon(coordinateE);
                        string convN = deglat(coordinateN);

                        richTextBox1.Text = richTextBox1.Text + "\r\n" + convN + " N " + convE + " E";

                        // }
                    }


                }
                MessageBox.Show("Airport coordinates loaded.");
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Simbolo 334020999N 0565330008E Tabas 2000#5

            string exportsimbolo=null;
            string[] simarr = coordinates.Split(';');
            int i = 1;
            foreach (string airport in simarr)
            {
                if (airport != "")
                {
                    string[] temp = airport.Split(',');
                    string temp1 = String.Format("Simbolo {0}000N {1}000E {2} 2000#{3}", temp[1], temp[2], temp[0], i.ToString()) + "\r\n";
                    i++;
                    exportsimbolo = exportsimbolo + temp1;
                }

           }
            richTextBox1.Text=exportsimbolo;
        }
    }
}