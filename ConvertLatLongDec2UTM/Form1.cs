using OfficeOpenXml;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace ConvertLatLongDec2UTM
{
    public partial class Form1 : Form
    {
        string coordinates=null;
        string coordinates2 = null;

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

        private void button3_Click(object sender, EventArgs e)
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
                        
                        string sampleout = "Linea {0}{1}{2}{3}N {4}{5}{6}{7}E {8}{9}{10}{11}N {12}{13}{14}{15}E";
                        string airportCode = firstSheet.Cells[row, 2].Value?.ToString().Trim();
                        string coordinateN = firstSheet.Cells[row, 4].Value?.ToString().Trim();
                        string coordinateE = firstSheet.Cells[row, 5].Value?.ToString().Trim();
                        string coordinateN2 = firstSheet.Cells[row, 6].Value?.ToString().Trim();
                        string coordinateE2 = firstSheet.Cells[row, 7].Value?.ToString().Trim();

                        coordinateN = deglat(coordinateN);
                        coordinateE = deglon(coordinateE);
                        coordinateN2 = deglat(coordinateN2);
                        coordinateE2 = deglon(coordinateE2);


                        string deg = coordinateN.Split("°")[0].Trim();
                        if (deg.Length==1) { deg = "0" + deg; }
                        string min = coordinateN.Split("°")[1].Split("\'")[0].Trim().Replace("\'","");
                        if (min.Length == 1) { 
                            
                            min = "0" + min; }

                        string sec = coordinateN.Split("°")[1].Split("\'")[1].Split(".")[0].Trim().Replace("\"", "");
                        if (sec.Length == 1) { sec = "0" + sec; }

                        string mil;
                        try
                        {
                            mil = coordinateN.Split("°")[1].Split("\'")[1].Split(".")[1].Trim().Replace("\"", "").Trim();
                            if (mil.Length == 1) { mil = "00" + mil; }
                            if (mil.Length == 2) { mil = "0" + mil; }
                            mil = mil.Substring(0, 3);
                        }
                        catch
                        {
                            mil = "000";
                        }

                        string deg2 = coordinateE.Split("°")[0].Trim();
                        if (deg2.Length == 1) { deg2 = "00" + deg2; }
                        if (deg2.Length == 2) { deg2 = "0" + deg2; }

                        string min2 = coordinateE.Split("°")[1].Split("\'")[0].Trim().Replace("\'", "");
                        if (min2.Length == 1)
                        {

                            min2 = "0" + min2;
                        }

                        string sec2 = coordinateE.Split("°")[1].Split("\'")[1].Split(".")[0].Trim().Replace("\"", "");
                        if (sec2.Length == 1) { sec2 = "0" + sec2; }

                        string mil2;
                        try
                        {
                            mil2 = coordinateE.Split("°")[1].Split("\'")[1].Split(".")[1].Trim().Replace("\"", "").Trim();
                            if (mil2.Length == 1) { mil2 = "00" + mil2; }
                            if (mil2.Length == 2) { mil2 = "0" + mil2; }
                            mil2 = mil2.Substring(0, 3);
                        }
                        catch
                        {
                            mil2 = "000";
                        }


                        string deg3 = coordinateN2.Split("°")[0].Trim();
                        if (deg3.Length == 1) { deg3 = "0" + deg3; }
                        string min3 = coordinateN2.Split("°")[1].Split("\'")[0].Trim().Replace("\'", "");
                        if (min3.Length == 1)
                        {

                            min3 = "0" + min3;
                        }

                        string sec3 = coordinateN2.Split("°")[1].Split("\'")[1].Split(".")[0].Trim().Replace("\"","");
                        if (sec3.Length == 1) { sec3 = "0" + sec3; }

                        string mil3;
                        try
                        {
                            mil3 = coordinateN2.Split("°")[1].Split("\'")[1].Split(".")[1].Trim().Replace("\"", "").Trim();
                            if (mil3.Length == 1) { mil3 = "00" + mil3; }
                            if (mil3.Length == 2) { mil3 = "0" + mil3; }
                            mil3 = mil3.Substring(0, 3);
                        }
                        catch
                        {
                           mil3 = "000";
                        }



                        string deg4 = coordinateE2.Split("°")[0].Trim();
                        if (deg4.Length == 1) { deg4 = "00" + deg4; }
                        if (deg4.Length == 2) { deg4 = "0" + deg4; }
                        string min4 = coordinateE2.Split("°")[1].Split("\'")[0].Trim().Replace("\'", "");
                        if (min4.Length == 1)
                        {

                            min4 = "0" + min4;
                        }

                        string sec4 = coordinateE2.Split("°")[1].Split("\'")[1].Split(".")[0].Trim().Replace("\"", "");
                        if (sec4.Length == 1) { sec4 = "0" + sec4; }

                        string mil4;
                        try
                        {
                            mil4 = coordinateE2.Split("°")[1].Split("\'")[1].Split(".")[1].Trim().Replace("\"", "").Trim();

                            if (mil4.Length == 1) { mil4 = "00" + mil4; }
                            if (mil4.Length == 2) { mil4 = "0" + mil4; }
                            mil4 = mil4.Substring(0, 3);
                        }
                        catch
                        {
                            mil4 = "000";
                        }

                        sampleout = String.Format(sampleout, deg, min, sec, mil, deg2, min2, sec2, mil2, deg3, min3, sec3, mil3, deg4, min4, sec4, mil4);



                        coordinates2 = coordinates2 + sampleout + "\r\n";





                    }


                }
            }
            richTextBox1.Text = coordinates2;

        }
    }
}