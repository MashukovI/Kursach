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
        string filePath = "D:\\Kursovaya\\Kursach\\NuRePrCalc\\ExcelDB.xlsx";
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

            this.Load += new EventHandler(Form1_Load);
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

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadInitialValues(); // Загружаем данные из Excel и отображаем в TextBox
        }

        private void LoadInitialValues()
        {

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

                    // Загружаем данные из Excel в TextBox при старте
                    textBoxT.Text = worksheet.Cells[3, 1].Value.ToString();
                    textBoxW.Text = worksheet.Cells[3, 6].Value.ToString();
                    textBoxD.Text = worksheet.Cells[3, 7].Value.ToString();
                    comboBoxA.Text = worksheet.Cells[3, 11].Value.ToString();
                    textBoxCO2.Text = worksheet.Cells[3, 2].Value.ToString();
                    textBoxH2O.Text = worksheet.Cells[3, 3].Value.ToString();
                    textBoxN2.Text = worksheet.Cells[3, 4].Value.ToString();
                    textBoxO2.Text = worksheet.Cells[3, 5].Value.ToString();

                    textBoxPr.Text = worksheet.Cells[43, 6].Value.ToString(); // F43
                    textBoxRe.Text = worksheet.Cells[8, 3].Value.ToString();  // C8

                    // Использование сопоставления для получения адреса выбранной ячейки
                    string cellAddress = cellMappings[selectedDescription];
                    textBoxNu.Text = worksheet.Range[cellAddress].Value.ToString();

                    string cellAddress2 = cellMappings2[selectedDescription];
                    textBoxV.Text = worksheet.Range[cellAddress2].Value.ToString();

                    textBoxAA.Text = worksheet.Cells[3, 13].Value.ToString(); // M3

                    workbook.Application.Calculate();
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
            }
            else
            {
                MessageBox.Show("Файл не найден: " + filePath);
            }
        }

        private bool ValidateSum()
        {
            try
            {
                // Преобразуем значения из текстовых полей в числа
                double co2 = double.Parse(textBoxCO2.Text);
                double h2o = double.Parse(textBoxH2O.Text);
                double n2 = double.Parse(textBoxN2.Text);
                double o2 = double.Parse(textBoxO2.Text);

                // Считаем сумму
                double sum = co2 + h2o + n2 + o2;

                // Проверяем, равна ли сумма 100
                if (sum != 100)
                {
                    // Если не равна, выводим ошибку
                    MessageBox.Show($"Ошибка: Сумма должна быть равна 100! Сейчас сумма = {sum}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false; // Возвращаем false, если сумма не равна 100
                }

                return true; // Возвращаем true, если сумма верна
            }
            catch (FormatException)
            {
                // Если введены некорректные данные (не числа)
                MessageBox.Show("Ошибка: Введите числовые значения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
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

            if (!ValidateSum())
            {
                return; // Если сумма неправильная, не продолжаем расчет
            }

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