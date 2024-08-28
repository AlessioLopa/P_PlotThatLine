namespace P_PlotThatLine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            double[] dataX = { 1, 2, 3, 4, 5 };
            double[] dataY = { 1, 4, 9, 16, 25 };

            formsPlot1.Plot.Add.Scatter(dataX, dataY);
            formsPlot1.Refresh();

        }

        private void formsPlot1_Load(object sender, EventArgs e)
        {

        }
    }
}
