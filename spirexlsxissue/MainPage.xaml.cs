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
        }

        void Button_Clicked(System.Object sender, System.EventArgs e)
        {
            // this code can be used to reproduce two issues I have found
            // 1) an exception being thrown when setting the cell color
            // 2) an exception being thrown when saving a chartsheet as a PDF

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

            // generate random data for the graph
            var rawDataWorksheet = workbook.Worksheets[1];
            var rand = new Random();
            for (int i = 0; i < 1000; i++)
            {
                rawDataWorksheet[i + 1, 1].NumberValue = i;
                rawDataWorksheet[i + 1, 1].NumberValue = rand.NextDouble();
            }

            // assign data to first graph sheet (comment this block of code to let the code run through without exception)
            var chart = workbook.Chartsheets[0];
            chart.ChartTitle = "test";
            chart.ChartType = ExcelChartType.ScatterLine;
            chart.DataRange = rawDataWorksheet.Range["B1:B1001"];
            chart.SeriesDataFromRange = false;
            chart.PrimaryCategoryAxis.Title = "Time";
            chart.PrimaryValueAxis.Title = "Value";
            var series = chart.Series.Add(ExcelChartType.ScatterLine);
            series.Name = "Oxygen";
            series.CategoryLabels = rawDataWorksheet.Range["A1:A1001"];
            series.Values = rawDataWorksheet.Range["B1:B1001"];

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
