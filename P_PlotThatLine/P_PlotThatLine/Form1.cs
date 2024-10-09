using Aspose.Cells;
using Aspose.Cells.Charts;
using ScottPlot;
using ScottPlot.Colormaps;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using Binance.Net;
using Binance.Net.Clients;
using System.Net;
using ScottPlot.Plottables;
using System.Drawing;

namespace P_PlotThatLine
{
    public partial class Form1 : Form
    {

       
        public Form1()
        {

            InitializeComponent();



            formsPlot1.Plot.Axes.DateTimeTicksBottom();
           

            var BTCUSDT = importBinanceData("BTCUSDT");
            var Plot_BTC = formsPlot1.Plot.Add.ScatterLine(BTCUSDT.date, BTCUSDT.price);
            Plot_BTC.LegendText = BTCUSDT.symbole;

            var ETHUSDT = importBinanceData("ETHUSDT");
            var Plot_ETH = formsPlot1.Plot.Add.ScatterLine(ETHUSDT.date, ETHUSDT.price);
            Plot_ETH.LegendText = ETHUSDT.symbole;


            var SOLUSDT = importBinanceData("SOLUSDT");
            var Plot_SOL = formsPlot1.Plot.Add.ScatterLine(SOLUSDT.date, SOLUSDT.price);
            Plot_ETH.LegendText = SOLUSDT.symbole;

            var MyCrosshair = formsPlot1.Plot.Add.Crosshair(0, 0);
            MyCrosshair.IsVisible = false;
            MyCrosshair.MarkerShape = MarkerShape.OpenCircle;
            MyCrosshair.MarkerSize = 15;

            ScottPlot.Plottables.Text MyHighlightText = formsPlot1.Plot.Add.Text("", 0,0);


            formsPlot1.MouseMove += (s, e) =>
            {

                // determine where the mouse is and get the nearest point
                Pixel mousePixel = new(e.Location.X, e.Location.Y);
                Coordinates mouseLocation = formsPlot1.Plot.GetCoordinates(mousePixel);

                var nearest = Plot_BTC.Data.GetNearest(mouseLocation, formsPlot1.Plot.LastRender);


                // place the crosshair over the highlighted point
                if (nearest.IsReal)
                {
                    MyCrosshair.IsVisible = true;
                    MyCrosshair.Position = nearest.Coordinates;

                    formsPlot1.Refresh();


                    MyHighlightText.Location = nearest.Coordinates;
                    MyHighlightText.LabelText = $"{nearest.X:0.##}, {nearest.Y:0.##}";
                    MyHighlightText.IsVisible = true;


                }

                // hide the crosshair when no point is selected
                if (!nearest.IsReal && MyCrosshair.IsVisible && MyHighlightText.IsVisible)
                {
      

                    MyCrosshair.IsVisible = false;
                    MyHighlightText.IsVisible = false;

                    formsPlot1.Refresh();
                    Text = $"No point selected";
                }
            };

        }

        public List<Cryptocurrency> importCryptoFromCSV(string path, string name)
        {
            // Import le fichier excel
            Workbook wb = new Workbook(path);

            // Récupère la première fiche
            Worksheet sheet = wb.Worksheets[0];

            int rows = sheet.Cells.MaxDataRow;
            int cols = sheet.Cells.MaxColumn;



            var data = new List<Cryptocurrency>();

            for (int x = 0 + 1; x < rows; x++)
            {
                int y = 0;

                data.Add(new Cryptocurrency(
                    name,
                    sheet.Cells[x, y++].StringValue,
                    sheet.Cells[x, y++].StringValue,
                    sheet.Cells[x, y++].StringValue,
                    sheet.Cells[x, y++].StringValue,
                    sheet.Cells[x, y++].StringValue,
                    sheet.Cells[x, y++].StringValue
                    ));

                y = 0;

            }

            return data;
        }

        public (List<decimal> price, List<DateTime> date, string symbole) importBinanceData(string symbol)
        {
            // Recupère les donnée historique depuis l'API Binance avec un interval de un jour
            var client = new BinanceRestClient();
            var klines = client.SpotApi.ExchangeData.GetKlinesAsync(symbol, Binance.Net.Enums.KlineInterval.OneDay);

            // Ajoute les données dans la liste correspondant à la date et au prix
            var price = klines.Result.Data.Select(item => item.OpenPrice).ToList();
            var date = klines.Result.Data.Select(item => item.OpenTime).ToList();

            return (price, date, symbol);

            // Affiche le graphe
            //var plot = formsPlot1.Plot.Add.ScatterLine(date, price);
            //plot.LegendText = symbol;

            //formsPlot1.Refresh();

        }

        private void formsPlot1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void PlotfilterDate(DateTime startDate, DateTime endDate)
        {
            formsPlot1.Reset();

            formsPlot1.Plot.Axes.DateTimeTicksBottom();

            // Importe les données depuis les fichier Excel
            var BTC = importCryptoFromCSV("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\bitcoin_2010-01-01_2024-08-28.xlsx", "BTCUSDT");
            var ETH = importCryptoFromCSV("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\ethereum_2015-07-30_2024-09-18.xlsx", "ETHUSDT");
            var SOL = importCryptoFromCSV("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\solana_2020-03-16_2024-09-25.xlsx", "SOLUSDT");

            // Sélectionne la date pour X et le prix d'ouverture pour Y (filtré selon les dates données)
            var plot1 = formsPlot1.Plot.Add.ScatterLine(
                BTC.Where(item => item.start > startDate && item.start < endDate)
                    .Select(item => item.start).ToArray(),
                BTC.Where(item => item.start > startDate && item.start < endDate)
                    .Select(item => item.open).ToArray());

            plot1.LegendText = BTC.First().name;

            var plot2 = formsPlot1.Plot.Add.ScatterLine(
                ETH.Where(item => item.start > startDate && item.start < endDate)
                    .Select(item => item.start).ToArray(),
                ETH.Where(item => item.start > startDate && item.start < endDate)
                    .Select(item => item.open).ToArray());

            plot2.LegendText = ETH.First().name;


            var plot3 = formsPlot1.Plot.Add.ScatterLine(
                SOL.Where(item => item.start > startDate && item.start < endDate)
                    .Select(item => item.start).ToArray(),
                SOL.Where(item => item.start > startDate && item.start < endDate)
                    .Select(item => item.open).ToArray());

            plot3.LegendText = SOL.First().name;


            formsPlot1.Refresh();


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            PlotfilterDate(dateTimePicker1.Value, dateTimePicker2.Value);

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        
    }
}
