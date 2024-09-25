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

        private List<SOL> importCSVDataSOL()
        {
            // Import le fichier excel
            Workbook wb = new Workbook("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\solana_2020-03-16_2024-09-25.xlsx");

            // Récupère la fiche des données
            Worksheet sheet = wb.Worksheets[0];

            int rows = sheet.Cells.MaxDataRow;
            int cols = sheet.Cells.MaxColumn;

            var data = new List<SOL>();

            for (int x = 0 + 1; x < rows; x++)
            {
                int y = 0;

                string start = sheet.Cells[x, y++].StringValue;
                string end = sheet.Cells[x, y++].StringValue;
                string open = sheet.Cells[x, y++].StringValue;
                string high = sheet.Cells[x, y++].StringValue;
                string low = sheet.Cells[x, y++].StringValue;
                string close = sheet.Cells[x, y++].StringValue;

                data.Add(new SOL(
                    start,
                    end,
                    open,
                    high,
                    low,
                    close
                    ));

                y = 0;

            }

            return data;
        }

        private List<ETH> importCSVDataETH()
        {
            // Import le fichier excel
            Workbook wb = new Workbook("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\ethereum_2015-07-30_2024-09-18.xlsx");

            // Récupère la fiche des données
            Worksheet sheet = wb.Worksheets[0];

            int rows = sheet.Cells.MaxDataRow;
            int cols = sheet.Cells.MaxColumn;

            var data = new List<ETH>();

            for (int x = 0 + 1; x < rows; x++)
            {
                int y = 0;

                string start = sheet.Cells[x, y++].StringValue;
                string end = sheet.Cells[x, y++].StringValue;
                string open = sheet.Cells[x, y++].StringValue;
                string high = sheet.Cells[x, y++].StringValue;
                string low = sheet.Cells[x, y++].StringValue;
                string close = sheet.Cells[x, y++].StringValue;

                data.Add(new ETH(
                    start,
                    end,
                    open,
                    high,
                    low,
                    close
                    ));

                y = 0;

            }

            return data;
        }

        private List<BTC> importCSVDataBTC()
        {
            // Import le fichier excel
            Workbook wb = new Workbook("C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\bitcoin_2010-01-01_2024-08-28.xlsx");

            // Récupère la fiche des données
            Worksheet sheet = wb.Worksheets[0];

            int rows = sheet.Cells.MaxDataRow;
            int cols = sheet.Cells.MaxColumn;

            var data = new List<BTC>();

            for (int x = 0 + 1; x < rows; x++)
            {
                int y = 0;

                string start = sheet.Cells[x, y++].StringValue;
                string end = sheet.Cells[x, y++].StringValue;
                string open = sheet.Cells[x, y++].StringValue;
                string high = sheet.Cells[x, y++].StringValue;
                string low = sheet.Cells[x, y++].StringValue;
                string close = sheet.Cells[x, y++].StringValue;

                data.Add(new BTC(
                    start,
                    end,
                    open,
                    high,
                    low,
                    close
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

            List<Decimal> price = new List<Decimal>();
            List<DateTime> date = new List<DateTime>();

            // Ajoute les données dans la liste correspondant à la date et au prix
            foreach (var item in klines.Result.Data)
            {
                price.Add(item.OpenPrice);
                date.Add(item.OpenTime);

            }

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
            var BTC = importCSVDataBTC();
            var ETH = importCSVDataETH();
            var SOL = importCSVDataSOL();

            // Filtrage des datas selon les dates em paramètre
            var filteredPriceBTC = BTC.Where(item => item.start > startDate && item.start < endDate).Select(item => item.open);
            var filteredDateBTC = BTC.Where(item => item.start > startDate && item.start < endDate).Select(item => item.start);

            var filteredPriceETH = ETH.Where(item => item.start > startDate && item.start < endDate).Select(item => item.open);
            var filteredDateETH = ETH.Where(item => item.start > startDate && item.start < endDate).Select(item => item.start);

            var filteredPriceSOL = SOL.Where(item => item.start > startDate && item.start < endDate).Select(item => item.open);
            var filteredDateSOL = SOL.Where(item => item.start > startDate && item.start < endDate).Select(item => item.start);

            // Affichage des graphiques filtrer
            formsPlot1.Plot.Add.ScatterLine(filteredDateBTC.ToArray(), filteredPriceBTC.ToArray());
            formsPlot1.Plot.Add.ScatterLine(filteredDateETH.ToArray(), filteredPriceETH.ToArray());
            formsPlot1.Plot.Add.ScatterLine(filteredDateSOL.ToArray(), filteredPriceSOL.ToArray());

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
