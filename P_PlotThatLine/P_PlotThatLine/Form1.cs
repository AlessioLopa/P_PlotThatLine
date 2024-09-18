using Aspose.Cells;
using Aspose.Cells.Charts;
using ScottPlot;
using ScottPlot.Colormaps;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;

namespace P_PlotThatLine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var data = importDataBTC();

            List<double> price = new List<double>();
            List<DateTime> date = new List<DateTime>();

            // Ajoute les données dans la liste correspondant à la date et au prix
            foreach (var item in data)
            {
                price.Add(item.open);
                date.Add(item.start);

            }

            formsPlot1.Plot.Add.ScatterLine(date, price);

            // Modifie la valeur 
            formsPlot1.Plot.Axes.DateTimeTicksBottom();
            formsPlot1.Refresh();


        }

        private List<BTC> importDataBTC()
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

        private void formsPlot1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            formsPlot1.Reset();

            DateTime startDate = dateTimePicker1.Value;
            DateTime endDate = dateTimePicker2.Value;

            var data = importDataBTC();


            List<double> filteredPrice = new List<double>();
            List<DateTime> filteredDate = new List<DateTime>();

            // Ajoute les données dans la liste correspondant à la date et au prix
            foreach (var item in data)
            {
                if (item.start > startDate && item.start < endDate)
                {
                    filteredPrice.Add(item.open);
                    filteredDate.Add(item.start);
                }
                
            }

            formsPlot1.Plot.Add.ScatterLine(filteredDate, filteredPrice);

            // Modifie la valeur 
            formsPlot1.Plot.Axes.DateTimeTicksBottom();
            formsPlot1.Refresh();

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
