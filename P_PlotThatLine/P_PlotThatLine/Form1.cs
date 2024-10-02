using Aspose.Cells;
using Aspose.Cells.Charts;
using ScottPlot;
using ScottPlot.Colormaps;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using Binance.Net;
using Binance.Net.Clients;
using System.Net;

namespace P_PlotThatLine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            formsPlot1.Plot.Axes.DateTimeTicksBottom();

            importBinanceData("BTCUSDT");
            importBinanceData("ETHUSDT");
            importBinanceData("SOLUSDT");
        }

        private List<Cryptocurrency> importCryptoFromCSV(string path, string name)
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

        public void importBinanceData(string symbol)
        {
            // Ràcupère les donnée historique depuis l'API Binance avec un interval de un jour
            var client = new BinanceRestClient();
            var klines = client.SpotApi.ExchangeData.GetKlinesAsync(symbol, Binance.Net.Enums.KlineInterval.OneDay);

            // Ajoute les données dans la liste correspondant à la date et au prix
            var price = klines.Result.Data.Select(item => item.OpenPrice).ToList();
            var date = klines.Result.Data.Select(item => item.OpenTime).ToList();
            
            formsPlot1.Plot.Add.ScatterLine(date, price);

            // Modifie la valeur 
            formsPlot1.Refresh();

        }

        public void dateFilter(DateTime startDate, DateTime endDate) {

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
            var BTC = importCryptoFromCSV("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\bitcoin_2010-01-01_2024-08-28.xlsx", "bitcoin");
            var ETH = importCryptoFromCSV("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\ethereum_2015-07-30_2024-09-18.xlsx", "Ethereum");
            var SOL = importCryptoFromCSV("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\solana_2020-03-16_2024-09-25.xlsx", "Solana");

            // Affichage des graphiques filtrer selon la date de début et de fin
            formsPlot1.Plot.Add.ScatterLine(BTC.Where(item => item.start > startDate && item.start < endDate).Select(item => item.start).ToArray(), BTC.Where(item => item.start > startDate && item.start < endDate).Select(item => item.open).ToArray());
            formsPlot1.Plot.Add.ScatterLine(ETH.Where(item => item.start > startDate && item.start < endDate).Select(item => item.start).ToArray(), ETH.Where(item => item.start > startDate && item.start < endDate).Select(item => item.open).ToArray());
            formsPlot1.Plot.Add.ScatterLine(SOL.Where(item => item.start > startDate && item.start < endDate).Select(item => item.start).ToArray(), SOL.Where(item => item.start > startDate && item.start < endDate).Select(item => item.open).ToArray());

            // Modifie la valeur 
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
    }
}
