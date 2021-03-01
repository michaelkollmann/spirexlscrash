using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Spire.Xls;
using Xamarin.Forms;

namespace spirexlsxissue
{
    public partial class MainPage : ContentPage
    {
        private const string PdfFileTemplateName = "pdfExportTemplate.xlsx";

        public MainPage()
        {
            InitializeComponent();
            Spire.License.LicenseProvider.SetLicenseKey("");

        }

        void Button_Clicked(System.Object sender, System.EventArgs e)
        {
            // this code can be used to reproduce two issues I have found
            // 1) an exception being thrown when setting the cell color
            // 2) an exception being thrown when saving a chartsheet as a PDF
            // 3) an exception being thrown when assigning more than 65535 data points to a graph

            // Get template file
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith(PdfFileTemplateName));

            // open template workbook
            using var excelFileTemplate = assembly.GetManifestResourceStream(resourceName);
            var workbook = new Workbook();
            workbook.LoadFromStream(excelFileTemplate, ExcelVersion.Version2016);
            workbook.Version = ExcelVersion.Version2016;

            // select the first worksheet and set the cell color
            var infoWorksheet = workbook.Worksheets[0];
            //infoWorksheet[1, 1].Style.Color = Color.Black;      // <-- 1) this throws an exception on iOS!

            // number of data rows
            // if this number is bigger than 65535 assigning the dataRange will crash
            var count = 100000;     // <-- 3) this causes an exception to be thrown if the unmber is set to more than 65535

            // generate random data for the graph
            var rawDataWorksheet = workbook.Worksheets[1];
            var rand = new Random();
            for (int i = 0; i < count; i++)
            {
                rawDataWorksheet[i + 1, 1].NumberValue = i;
                rawDataWorksheet[i + 1, 1].NumberValue = rand.NextDouble();
            }

            // assign data to first graph sheet (comment this block of code to let the code run through without exception)
            var chart = workbook.Chartsheets[0];
            chart.ChartTitle = "test";
            chart.ChartType = ExcelChartType.ScatterLine;
            //chart.DataRange = rawDataWorksheet.Range[string.Format("B1:B{0}", count + 1)];        // <-- 3) here the exception about the 65535 data points is being thrown
            chart.SeriesDataFromRange = false;
            chart.PrimaryCategoryAxis.Title = "Time";
            chart.PrimaryValueAxis.Title = "Value";
            var series = chart.Series.Add(ExcelChartType.ScatterLine);
            series.Name = "Oxygen";
            series.CategoryLabels = rawDataWorksheet.Range[string.Format("A1:A{0}", count + 1)];
            series.Values = rawDataWorksheet.Range[string.Format("B1:B{0}", count + 1)];

            // save as pdf
            var name = Path.Combine(GetBasePath(), "testfile.xlsx");
            if (File.Exists(name))
            {
                File.Delete(name);
            }
            workbook.SaveToFile(name, FileFormat.PDF);      // <-- 2) this throws an exception. The exception is not being thrown if I do not assign data to the chartsheet above
        }

        // get a temporary folder to save our file to
        public string GetBasePath()
        {
            switch (Device.RuntimePlatform)
            {
                case Device.Android:
                    return Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                case Device.iOS:
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "..", "Library");
                default:
                    throw new NotImplementedException("Platform not supported");
            }
        }
    }
}
