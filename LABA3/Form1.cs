using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace LABA3
{
    public partial class Form1 : Form
    {
		private readonly string path = @"C:\Users\Honor\Desktop\prog\LABA3\LABA3\Lab3.1.xlsm";

		private Excel.Application excel1;
		private Excel.Workbook workBook1;
		public Form1()
		{
			InitializeComponent();
			try
			{
				Excel.Application excel = new Excel.Application { Visible = true };
				Excel.Workbook workBook = excel.Workbooks.Open(path);
				Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];

				int funcsCount = (int)workSheet.Cells[1, "F"].Value;
				for (int i = 1; i <= funcsCount; i++)
					comboBox1.Items.Add(workSheet.Cells[i, "A"].Value);

				excel1 = excel;
				workBook1 = workBook;
			}
			catch (Exception er)
			{
				MessageBox.Show("Error", er.Message);
				Close();
			}
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			Excel.Worksheet workSheet = (Excel.Worksheet)workBook1.Sheets[2];
			workSheet.Cells[2, "F"].Value = (sender as ComboBox).SelectedIndex + 1;
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Excel.Worksheet workSheet2 = (Excel.Worksheet)workBook1.Sheets[2];
			Excel.Worksheet workSheet3 = (Excel.Worksheet)workBook1.Sheets[3];

			int rowIndex = 0;

			for (int x = 0; x <= 10; x++)
			{
				workSheet2.Cells[3, "F"].Value = x;
				double y = workSheet2.Cells[6, "F"].Value;

				workSheet3.Cells[rowIndex + 1, "H"].Value = x;
				workSheet3.Cells[rowIndex + 1, "I"].Value = y;

				rowIndex++;
			}

			Excel.ChartObjects chartObjs = (Excel.ChartObjects)workSheet3.ChartObjects(Type.Missing);
			Excel.ChartObject myChart = chartObjs.Add(20, 60, 200, 200);
			Excel.Chart chart = myChart.Chart;
			Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
			Excel.Series series = seriesCollection.NewSeries();
			series.XValues = workSheet3.get_Range("H1", "H10");
			series.Values = workSheet3.get_Range("I1", "I10");
			chart.ChartType = Excel.XlChartType.xlXYScatterSmooth;
		}
	}
}
