using System;
using System.Collections.Generic;
using Spire.Xls;
using Xamarin.Forms;

namespace spirexlsxissue
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }

        void Button_Clicked(System.Object sender, System.EventArgs e)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets.Create();

            worksheet[1, 1].Style.Color = Color.Black;
        }
    }
}
