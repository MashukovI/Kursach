using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;

namespace NuRePrCalc
{
    public partial class Form1 : Form
    {

        private Dictionary<string, string> cellMappings = new Dictionary<string, string>
        {
            { "Шахматное", "D8" },
            { "Коридорное", "E8" }

        };

        private Dictionary<string, string> cellMappings2 = new Dictionary<string, string>
        {
            { "Шахматное", "G8" },
            { "Коридорное", "H8" }

        };

        private Dictionary<string, string> imageMappings = new Dictionary<string, string>
        {
            { "Шахматное", "Shah.png" },
            { "Коридорное", "Kor.png" }
        };

        private Dictionary<string, string> categoryRangeMappings = new Dictionary<string, string>
        {
            { "Шахматное", "R12:R20" }, // Диапазон для "Шахматное"
            { "Коридорное", "S12:S20" } // Диапазон для "Коридорное"
        };

        public Form1()
        {


            InitializeComponent();

            for (int i = 10; i <= 90; i += 10)
            {
                comboBoxA.Items.Add(i.ToString());
            }

            foreach (var key in cellMappings.Keys)
            {
                comboBoxNuCells.Items.Add(key);
            }
            comboBoxNuCells.SelectedIndex = 0;

            LoadImage();
        }

        private void comboBoxNuCells_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadImage(); // Загружаем изображение при изменении выбора
        }

        private void LoadImage()
        {
        
            string selectedDescription = comboBoxNuCells.SelectedItem.ToString();
            string imagePath = imageMappings[selectedDescription];

            pictureBoxNu.ImageLocation = imagePath;
            pictureBoxNu.Load(); 
        }

        private void LoadExcelDataAndPlot()
        {
            string filePath = "D:\\Kursovaya\\NuRePrCalc\\ExcelDB.xlsx";

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                string selectedDescription = comboBoxNuCells.SelectedItem.ToString();
                string categoryRangeAddress = categoryRangeMappings[selectedDescription];
                var categoryRange = worksheet.Range(categoryRangeAddress);

                var valueRange = worksheet.Range("N12:N20");

                double[] valueArray = new double[valueRange.RowCount()];

                Series series = new Series("Зависимость");
                series.ChartType = SeriesChartType.Line;
                series.Color = Color.Blue;  // Синий цвет линии
                series.BorderWidth = 3;  // Толщина линии
                series.BorderDashStyle = ChartDashStyle.Solid;
                series.MarkerStyle = MarkerStyle.Circle;
                series.MarkerSize = 8;
                series.MarkerColor = Color.Red;


                for (int i = 1; i <= valueRange.RowCount(); i++)
                {
                    double value = valueRange.Cell(i, 1).GetValue<double>();
                    valueArray[i - 1] = Math.Round(value, 2);  // Округляем до 2 знаков
                }

                for (int i = 1; i <= categoryRange.RowCount(); i++)
                {
                    double category = Math.Round(categoryRange.Cell(i, 1).GetValue<double>(),2);
                    double value = Math.Round(valueRange.Cell(i, 1).GetValue<double>(), 2);

                    series.Points.AddXY(category, value);
                }
                if (chartVA.Series.Count == 0)
                {
                    chartVA.Series.Add(series);
                }


                chartVA.ChartAreas[0].AxisX.Title = "Конвективный коэффициент, Вт/(м²·К)";
                chartVA.ChartAreas[0].AxisX.TitleFont = new Font("Arial", 12, FontStyle.Bold);
                chartVA.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 10);
                chartVA.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray;  // Цвет сетки

                // Настройка оси Y
                chartVA.ChartAreas[0].AxisY.Title = "Угол атаки, °";
                chartVA.ChartAreas[0].AxisY.TitleFont = new Font("Arial", 12, FontStyle.Bold);
                chartVA.ChartAreas[0].AxisY.LabelStyle.Font = new Font("Arial", 10);
                chartVA.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray;  // Цвет сетки

                // Установка формата для осей (округление до 2 знаков)
                chartVA.ChartAreas[0].AxisX.LabelStyle.Format = "N2";
                chartVA.ChartAreas[0].AxisY.LabelStyle.Format = "N2";

                // Настройка сетки на графике
                chartVA.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                chartVA.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                // Добавляем легенду
                chartVA.Legends.Clear();
                chartVA.Legends.Add(new Legend("Legend"));
                chartVA.Legends[0].Docking = Docking.Top;
                chartVA.Legends[0].Font = new Font("Arial", 10, FontStyle.Bold);
                chartVA.Legends[0].ForeColor = Color.DarkBlue;

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string filePath = "D:\\Kursovaya\\NuRePrCalc\\ExcelDB.xlsx";

            string selectedDescription = comboBoxNuCells.SelectedItem.ToString();

            if (System.IO.File.Exists(filePath))
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excelApp = Activator.CreateInstance(excelType);

                try
                {
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    dynamic workbook = excelApp.Workbooks.Open(filePath);
                    dynamic worksheet = workbook.Sheets[1];

                    worksheet.Cells[3, 1].Value = textBoxT.Text;
                    worksheet.Cells[3, 6].Value = textBoxW.Text;
                    worksheet.Cells[3, 7].Value = textBoxD.Text;
                    worksheet.Cells[3, 11].Value = Int32.Parse(comboBoxA.Text);
                    worksheet.Cells[3, 2].Value = textBoxCO2.Text;
                    worksheet.Cells[3, 3].Value = textBoxH2O.Text;
                    worksheet.Cells[3, 4].Value = textBoxN2.Text;
                    worksheet.Cells[3, 5].Value = textBoxO2.Text;

                    workbook.Application.Calculate();

                    textBoxPr.Text = worksheet.Cells[43, 6].Value.ToString(); // F43
                    textBoxRe.Text = worksheet.Cells[8, 3].Value.ToString();  // C8

                    // Использование сопоставления для получения адреса выбранной ячейки
                    string cellAddress = cellMappings[selectedDescription];
                    textBoxNu.Text = worksheet.Range[cellAddress].Value.ToString();

                    string cellAddress2 = cellMappings2[selectedDescription];
                    textBoxV.Text = worksheet.Range[cellAddress2].Value.ToString();

                    textBoxAA.Text = worksheet.Cells[3, 13].Value.ToString(); // M3

                    workbook.Save();
                    workbook.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
                finally
                {
                    Marshal.ReleaseComObject(excelApp);
                }


                chartVA.Series.Clear();
                chartVA.ChartAreas.Clear();
                chartVA.Legends.Clear();

                if (chartVA.ChartAreas.Count == 0) // Проверяем, есть ли уже области
                {
                    chartVA.ChartAreas.Add("ChartArea1");
                }

                LoadExcelDataAndPlot();
            }
            else
            {
                MessageBox.Show("Файл не найден: " + filePath);
            }
        }
    }
}