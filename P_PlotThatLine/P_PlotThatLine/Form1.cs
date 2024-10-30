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
using Aspose.Cells.Drawing;

namespace P_PlotThatLine
{
    public partial class Form1 : Form
    {

       
        public Form1()
        {

            InitializeComponent();


            // Initialise l'axe X avec les dates
            formsPlot1.Plot.Axes.DateTimeTicksBottom();

            List<ScottPlot.Plottables.Scatter> MyScatters = new();

            // Importe les données de binance, affiche le graphique 
            var BTCUSDT = importBinanceData("BTCUSDT");
            var Plot_BTC = formsPlot1.Plot.Add.ScatterLine(BTCUSDT.date, BTCUSDT.price);
            Plot_BTC.LegendText = BTCUSDT.symbole;

            var ETHUSDT = importBinanceData("ETHUSDT");
            var Plot_ETH = formsPlot1.Plot.Add.ScatterLine(ETHUSDT.date, ETHUSDT.price);
            Plot_ETH.LegendText = ETHUSDT.symbole;


            var SOLUSDT = importBinanceData("SOLUSDT");
            var Plot_SOL = formsPlot1.Plot.Add.ScatterLine(SOLUSDT.date, SOLUSDT.price);
            Plot_ETH.LegendText = SOLUSDT.symbole;

            MyScatters.Add(Plot_BTC);
            MyScatters.Add(Plot_ETH);
            MyScatters.Add(Plot_SOL);

            // Créer un crosshair et ne l'affiche pas et un marker
            var MyCrosshair = formsPlot1.Plot.Add.Crosshair(0, 0);
            ScottPlot.Plottables.Marker marker = formsPlot1.Plot.Add.Marker(0, 0);
            MyCrosshair.IsVisible = false;
            marker.Shape = MarkerShape.OpenCircle;
            marker.Size = 15;
            marker.LineWidth = 2;


            // Créer un texte qui n'affiche rien pour les détails
            ScottPlot.Plottables.Text detailsText = formsPlot1.Plot.Add.Text("", 0,0);
            detailsText.LabelAlignment = Alignment.LowerLeft;
            detailsText.LabelBold = true;
            detailsText.OffsetX = 7;
            detailsText.OffsetY = -7;

            formsPlot1.Refresh();

            formsPlot1.MouseMove += (s, e) =>
            {

                // Détermine où est le curseur
                Pixel mousePixel = new(e.Location.X, e.Location.Y);
                Coordinates mouseLocation = formsPlot1.Plot.GetCoordinates(mousePixel);
                
                Dictionary<int, DataPoint> nearestPoints = new();

                // Récupère les pointes les plus près du curseur pour chaque graphique, puis les ajoutes à une liste
                nearestPoints = MyScatters.Select((scatter, index) => new
                {
                    Index = index,
                    NearestPoint = scatter.Data.GetNearest(mouseLocation, formsPlot1.Plot.LastRender)
                }).ToDictionary(item => item.Index, item => item.NearestPoint);

                bool pointSelected = false;
                int scatterIndex = -1;
                double smallestDistance = double.MaxValue;

                //
                // TODO passer en LINQ
                for (int i = 0; i < nearestPoints.Count; i++)
                {
                    if (nearestPoints[i].IsReal)
                    {
                        // Calcule la distance entre le point et le curseur
                        double distance = nearestPoints[i].Coordinates.Distance(mouseLocation);
                        if (distance < smallestDistance)
                        {
                            // Ajoute l'index
                            scatterIndex = i;
                            pointSelected = true;
                            // Met à jour la bonne distance
                            smallestDistance = distance;
                        }
                    }
                }

                // Affiche les détails si le point est sélectionner
                if (pointSelected)
                {
                    ScottPlot.Plottables.Scatter scatter = MyScatters[scatterIndex];
                    DataPoint point = nearestPoints[scatterIndex];

                    MyCrosshair.IsVisible = true;
                    MyCrosshair.Position = point.Coordinates;
                    MyCrosshair.LineColor = scatter.MarkerStyle.FillColor;

                    marker.IsVisible = true;
                    marker.Location = point.Coordinates;
                    marker.MarkerStyle.LineColor = scatter.MarkerStyle.FillColor;

                    detailsText.IsVisible = true;
                    detailsText.Location = point.Coordinates;
                    // TODO changer la valeur X par la bonne date
                    detailsText.LabelText = $"{point.X}, {point.Y:0.##}";
                    detailsText.LabelFontColor = scatter.MarkerStyle.FillColor;

                    formsPlot1.Refresh();
                    base.Text = $"Selected Scatter={scatter.LegendText}, Index={point.Index}, X={point.X:0.##}, Y={point.Y:0.##}";
                }

                // Cache le crosshaire, le text et le text si aucun point n'est séléctionner
                if (!pointSelected && MyCrosshair.IsVisible)
                {
                    MyCrosshair.IsVisible = false;
                    marker.IsVisible = false;
                    detailsText.IsVisible = false;
                    formsPlot1.Refresh();
                    base.Text = $"No point selected";
                }

                
            };

        }

        /// <summary>
        /// Import les données d'un fichier csv et l'ajoute à une classe avec le nom correspondant
        /// </summary>
        /// <param name="path">Correspond au chemin du fichier csv</param>
        /// <param name="name">Correspond au nom du symbole correspondant au fichier (Exemple : BTC, ETH, etc...</param>
        /// <returns></returns>
        public List<Cryptocurrency> importCryptoFromCSV(string path, string name)
        {
            // Import le fichier excel
            Workbook wb = new Workbook(path);

            // Récupère la première fiche
            Worksheet sheet = wb.Worksheets[0];

            int rows = sheet.Cells.MaxDataRow;
            int cols = sheet.Cells.MaxColumn;

            // Créer une nouvelle classe et ajoute les données
            var data = new List<Cryptocurrency>();

            // (Pas possible de changer en LINQ)
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

        /// <summary>
        /// Importe les données depuis binance en utilisant son API
        /// </summary>
        /// <param name="symbol">Correspond au symbole (nom) de la cryptomonaie (Exemple : BTCUSD, etc...)</param>
        /// <returns></returns>
        public (List<decimal> price, List<DateTime> date, string symbole) importBinanceData(string symbol)
        {
            // Recupère les donnée historique depuis l'API Binance avec un interval de un jour
            var client = new BinanceRestClient();
            var klines = client.SpotApi.ExchangeData.GetKlinesAsync(symbol, Binance.Net.Enums.KlineInterval.OneDay);

            // Ajoute les données dans la liste correspondant à la date et au prix
            var price = klines.Result.Data.Select(item => item.OpenPrice).ToList();
            var date = klines.Result.Data.Select(item => item.OpenTime).ToList();

            return (price, date, symbol);

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

        /// <summary>
        /// Filtre les données avec les dates
        /// </summary>
        /// <param name="startDate">Date de début</param>
        /// <param name="endDate">Date de fin</param>
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
