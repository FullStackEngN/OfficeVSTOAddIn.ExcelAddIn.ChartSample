/// <summary>
/// WARNING: ANY USE BY YOU OF THE SAMPLE CODE PROVIDED IN THIS FILE IS AT YOUR OWN RISK. 
/// Microsoft provides this code "as is" without warranty of any kind, either express or implied, 
/// including but not limited to the implied warranties of merchantability and/or fitness for a particular purpose.
/// </summary>

using System;
using System.Drawing;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.ChartSample
{
    public partial class RibbonCharts
    {
        private void RibbonCharts_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonAddChart_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
        Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            Excel.Range cells = worksheet.Range["A1", "D4"];

            // Create a first chart
            double left = 10.00;
            double top = 10.00;
            double width = 200.00;
            double height = 200.00;

            Chart chart1 = worksheet.Controls.AddChart(left, top, width, height, "test" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
            chart1.ChartType = Excel.XlChartType.xlColumnClustered;
            chart1.SetSourceData(cells);

            chart1.PlotArea.Border.Color = ColorTranslator.ToOle(Color.Green);

            // Create a new chart2
            left = 220.00;
            top = 10.00;
            width = 200.00;
            height = 200.00;

            Chart chart2 = worksheet.Controls.AddChart(left, top, width, height, "test" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
            chart2.ChartType = Excel.XlChartType.xlColumnClustered;
            chart2.SetSourceData(cells);

            chart2.PlotArea.Border.Color = ColorTranslator.ToOle(Color.Red);

            // We must select this PlotArea, otherwise we will see exception
            // Exception message: System.Runtime.InteropServices.COMException
            // HResult = 0x80004005
            // Message = Unspecified error(Exception from HRESULT: 0x80004005(E_FAIL))
            // Source =< Cannot evaluate the exception source>
            // StackTrace:
            // < Cannot evaluate the exception stack trace >
            chart2.PlotArea.Select();

            // If we don't select PlotArea first
            // We can get its data first
            // Then set its data again
            // Then we won't see exception
            // var tempLeft = chart.PlotArea.Left;

            chart2.PlotArea.Left = 15.00;
            chart2.PlotArea.Top = 15.00;
            chart2.PlotArea.Width = 150.00;
            chart2.PlotArea.Height = 150.00;
        }

        private void SetSampleData(Worksheet worksheet)
        {
            worksheet.Range["A1", "A4"].Value2 = 1;
            worksheet.Range["B1", "B4"].Value2 = 2;
            worksheet.Range["C1", "C4"].Value2 = 3;
            worksheet.Range["D1", "D4"].Value2 = 4;
        }

        private void buttonAddCharts_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
        Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            Excel.Range cells = worksheet.Range["A1", "D4"];

            double chartLeft, chartTop, chartWidth, chartHeight;
            double plotLeft = 15.00;
            double plotTop = 15.00;
            double plotWidth = 150.00;
            double plotHeight = 150.00;

            int random = new Random().Next(1, 50);

            int row = 1;

            for (int i = 1; i < random; i++)
            {
                if (i % 2 != 0)
                {
                    row = i / 2 + 1;

                    chartLeft = 10.00;
                    chartTop = row * 220.00;
                    chartWidth = 200.00;
                    chartHeight = 200.00;

                    Chart leftChart = worksheet.Controls.AddChart(chartLeft, chartTop, chartWidth, chartHeight, "test" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                    leftChart.ChartType = Excel.XlChartType.xlColumnClustered;
                    leftChart.SetSourceData(cells);

                    leftChart.PlotArea.Border.Color = ColorTranslator.ToOle(Color.Green);
                }
                else
                {
                    row = i / 2;

                    chartLeft = 220.00;
                    chartTop = row * 220.00;
                    chartWidth = 200.00;
                    chartHeight = 200.00;

                    Chart rightChart = worksheet.Controls.AddChart(chartLeft, chartTop, chartWidth, chartHeight, "test" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                    rightChart.ChartType = Excel.XlChartType.xlColumnClustered;
                    rightChart.SetSourceData(cells);

                    rightChart.PlotArea.Border.Color = ColorTranslator.ToOle(Color.Red);

                    // We must select this PlotArea, otherwise we will see exception
                    // Exception message: System.Runtime.InteropServices.COMException
                    // HResult = 0x80004005
                    // Message = Unspecified error(Exception from HRESULT: 0x80004005(E_FAIL))
                    // Source =< Cannot evaluate the exception source>
                    // StackTrace:
                    // < Cannot evaluate the exception stack trace >
                    rightChart.PlotArea.Select();

                    rightChart.PlotArea.Left = plotLeft;
                    rightChart.PlotArea.Top = plotTop;
                    rightChart.PlotArea.Width = plotWidth;
                    rightChart.PlotArea.Height = plotHeight;
                }
            }
        }

        private void buttonSetData_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
  Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            SetSampleData(worksheet);
        }
    }
}
