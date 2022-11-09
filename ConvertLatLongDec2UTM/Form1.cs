using OfficeOpenXml;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace ConvertLatLongDec2UTM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


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
                        string convE =deglon(coordinateE);

                        // }
                    }


                }

            }
            
        }
    }
}