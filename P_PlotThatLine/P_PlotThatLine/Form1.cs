using Aspose.Cells;
using Aspose.Cells.Charts;
using ScottPlot;
using ScottPlot.Colormaps;
using System.Diagnostics;

namespace P_PlotThatLine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

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

        private void formsPlot1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
