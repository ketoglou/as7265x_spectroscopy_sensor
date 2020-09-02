using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Diagnostics;

namespace Spectroscopy_sensor
{
    public partial class Form1 : Form
    {

        private SerialPort serialPort = null;
        private BackgroundWorker backgroundWorker1;
        private Boolean sample_continue = false;
        private ChartArea chartArea1;
        private Series series1;
        private Series series2;
        private Chart chart;
        private Boolean firstTimeChart = true;
        private double Gamma = 0.80;
        private double IntensityMax = 255;
        private Mutex mut2 = new Mutex();
        private ToolTip tt = null;
        private Point tl = Point.Empty;
        private SaveFileDialog saveFileDialog1;

        public Form1()
        {
            InitializeComponent();
            
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;

            this.FormClosing += Form1_FormClosing;

            enable_configuration_panel(false);
            comboBox5.SelectedIndex = 1;
            button12.Hide();
            button12.Enabled = false;

            comboBox1.Click += ComboBox1_Click;
            // Get a list of serial port names.
            string[] ports = SerialPort.GetPortNames();
            // Add each COM to the combobox
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.DoWork += new DoWorkEventHandler(BackgroundWorker1_DoWork);
            backgroundWorker1.RunWorkerCompleted += BackgroundWorker1_RunWorkerCompleted;
            backgroundWorker1.WorkerSupportsCancellation = true;
            
            //Initialize chart 
            chartArea1 = new ChartArea();
            chartArea1.Name = "chartArea1";
            chartArea1.AxisX.Title = "Wavelength (nm)";
            chartArea1.AxisY.Title = "Frequency";
            chartArea1.AxisX.Maximum = 940;
            chartArea1.AxisX.Minimum = 410;
            chartArea1.AxisY.Minimum = 0;
            chartArea1.AxisX.MajorGrid.Enabled = false;
            chartArea1.AxisY.MajorGrid.Enabled = false;
            chartArea1.CursorX.LineColor = Color.Black;
            chartArea1.CursorY.LineColor = Color.Black;

            //Series 1
            series1 = new Series();
            series1.Name = "series1";
            series1.ChartArea = "chartArea1";
            series1.Color = Color.Black;
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.MarkerColor = Color.Black;
            series1.ChartType = SeriesChartType.SplineArea;
            

            //Series 2
            series2 = new Series();
            series2.Name = "series2";
            series2.ChartArea = "chartArea1";
            series2.Color = Color.Black;
            series2.MarkerStyle = MarkerStyle.Circle;
            series2.MarkerColor = Color.Black;
            series2.ChartType = SeriesChartType.Spline;

            //Chart
            chart = new Chart();
            chart.Name = "chart1";
            chart.Dock = DockStyle.Fill;
            chart.ChartAreas.Add(chartArea1);
            chart.Location = new Point(61, 109);
            chart.ChartAreas[0].AxisX.IsMarginVisible = true;
            chart.Size = new Size(300, 300);
            chart.TabIndex = 0;
            chart.Text = "chart1";
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            chart.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            chart.MouseWheel += Chart_MouseWheel;
            this.splitContainer1.Panel1.Controls.Add(chart);

            saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save chart";
            saveFileDialog1.DefaultExt = "png";
            saveFileDialog1.Filter = "png files (*.png)|*.png|jpeg files (*.jpg)|*.jpg|bmp files (*.bmp)|*.bmp|gif files (*.gif)|*.gif|tiff files (*.tiff)|*.tiff";
            saveFileDialog1.CheckPathExists = true;

        }

        //**************************************************************************************************************************
        //---------------------------------SERIAL PORT CONNECT FUNCTIONS------------------------------------------------------------
        private void ComboBox1_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            // Get a list of serial port names.
            string[] ports = SerialPort.GetPortNames();
            // Add each COM to the combobox
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            if (comboBox1.Items.Count == 0)
                comboBox1.Text = "";
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count == 0)
            {
                connect_info_label.Text = "No COMs found.";
                button4.Text = "Connect";
                if (serialPort != null)
                {
                    serialPort.Close();
                    serialPort = null;
                }
                enable_configuration_panel(false);
            }
            else
            {
                if (serialPort != null)
                {
                    if (sample_continue)
                        backgroundWorker1.CancelAsync();
                    sample_continue = false;
                    serialPort.Close();
                    serialPort = null;
                }
                if (button4.Text == "Disconnect")
                {
                    connect_info_label.Text = "Disconnected";
                    button4.Text = "Connect";
                    enable_configuration_panel(false);
                }
                else
                {
                    try
                    {
                        if (comboBox1.SelectedItem != null)
                        {
                            serialPort = new SerialPort(comboBox1.SelectedItem.ToString(), 115200, Parity.None, 8, StopBits.One);
                            serialPort.ReadTimeout = 100; //100msec read timeout
                            serialPort.Open();
                            button4.Text = "Disconnect";
                            connect_info_label.Text = "Connected";
                            //Initialize module
                            enable_configuration_panel(true);
                        }else
                            connect_info_label.Text = "Choose a COM.";
                    }
                    catch(Exception ex)
                    {
                        serialPort.Close();
                        if (ex is UnauthorizedAccessException)
                            connect_info_label.Text = "COM is open by another software.";
                        else if(ex is IOException)
                            connect_info_label.Text = "COM not found";
                        else if(ex is NullReferenceException)
                            connect_info_label.Text = "Choose a COM.";
                    }

                }
            }
        }

        //**************************************************************************************************************************
        //---------------------------------OTHER FUNCTIONS--------------------------------------------------------------------------

        string serial_write(string command)
        {
            if (serialPort != null)
            {
                try
                {
                    mut2.WaitOne();
                    serialPort.Write(command + "\n");
                    try
                    {
                        string respond = serialPort.ReadLine(); ;
                        mut2.ReleaseMutex();
                        return respond;
                    }
                    catch (TimeoutException)
                    {
                        serial_close();
                        connect_info_label.Text = "COM not respond, try change COM or reconnect.";
                        button4.Text = "Connect";
                    }
                    catch (IOException)
                    {
                        serial_close();
                    }
                }
                catch (InvalidOperationException)
                {
                    serial_close();
                    connect_info_label.Text = "COM not respond, try change COM or reconnect.";
                    button4.Text = "Connect";
                }
            }
            mut2.ReleaseMutex();
            return null;
        }

        void serial_close()
        {
            if (serialPort == null)
            {
                connect_info_label.Text = "";
            }
            else
            {
                mut2.WaitOne();
                serialPort.Close();
                serialPort = null;
                mut2.ReleaseMutex();
                MessageBox.Show("Connection lost,check connections!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            enable_configuration_panel(false);
            button4.Text = "Connect";
        }

        void enable_configuration_panel(Boolean set)
        {
            if (set)
            {   
                try
                {
                    string[] respond = serial_write("ATVERSW").Split(new string[] { "OK" }, StringSplitOptions.None);
                    if (respond != null)
                        fw_version.Text = "Firmware Version : " + respond[0];
                    respond = serial_write("ATVERHW").Split(new string[] { "OK" }, StringSplitOptions.None);
                    if (respond != null)
                        hw_version.Text = "Hardware Version : " + respond[0];
                    respond = serial_write("ATTEMP").Split(new string[] { "OK" }, StringSplitOptions.None);
                    if (respond != null)
                        temperature_label.Text = "Device Temperature : " + respond[0];
                    respond = serial_write("ATGAIN").Split(new string[] { "OK" }, StringSplitOptions.None);
                    if (respond != null)
                    {
                        comboBox5.SelectedIndex = Int32.Parse(respond[0]);
                        label6.Text = "Gain Set : " + comboBox5.SelectedItem.ToString();
                    }
                    if (serial_write("ATINTTIME=64") != "OK")
                        serial_close();

                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button5.Enabled = true;
                    button6.Enabled = true;
                    button7.Enabled = true;
                    button8.Enabled = true;
                    button9.Enabled = true;
                    button10.Enabled = true;
                    button11.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                    comboBox6.Enabled = true;
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton2.Select();
                    button9.Text = "Start Sampling";
                }
                catch(Exception)
                {
                    serial_close();
                }
              
            }
            else
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                button7.Enabled = false;
                button8.Enabled = false;
                button9.Enabled = false;
                button10.Enabled = false;
                button11.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                comboBox4.Enabled = false;
                comboBox5.Enabled = false;
                comboBox6.Enabled = false;
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                button9.Text = "Start Sampling";
            }
        }

        //**************************************************************************************************************************
        //---------------------------------LED FUNCTIONS----------------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED1=1");      
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED1=0");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLEDC=0x"+comboBox2.SelectedIndex.ToString()+ comboBox6.SelectedIndex.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED3=1");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED3=0");
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLEDD=0x" + comboBox3.SelectedIndex.ToString() + "0");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED5=1");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED5=0");     
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLEDE=0x" + comboBox4.SelectedIndex.ToString()+"0");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED0=1");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLED0=0");
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serialPort != null)
                serial_write("ATLEDC=0x" + comboBox2.SelectedIndex.ToString() + comboBox6.SelectedIndex.ToString());
        }
        //**************************************************************************************************************************
        //---------------------------------PANEL FUNCTIONS--------------------------------------------------------------------------
        private void button5_Click(object sender, EventArgs e)
        {
            if (serialPort != null)
            {
                string respond = serial_write("ATRST");
                if (respond == "ERROR")
                    MessageBox.Show("Software reset successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(serialPort != null)
            {
                if (serial_write("ATGAIN="+comboBox5.SelectedIndex.ToString()) == "OK")
                    label6.Text = "Gain set : " + comboBox5.SelectedItem.ToString();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            sample_continue = !sample_continue;
            if (sample_continue)
            {
                button9.Text = "Stop Sampling";
                button12.Show();
                button12.Enabled = true;
                if (firstTimeChart)
                {
                    chart.MouseMove += Chart_MouseMove;
                    firstTimeChart = false;
                }
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                backgroundWorker1.CancelAsync();
                button9.Text = "Start Sampling";
            }
        }

        //**************************************************************************************************************************
        //---------------------------------THREADS----------------------------------------------------------------------------------

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] raw_data_st;
            int[] raw_data;
            float[] calibrated_data;
            int[] wavelenghts = { 610, 680, 730, 760, 810, 860, 560, 585, 645, 705, 900, 940, 410, 435, 460, 485, 510, 535 };

            if (serialPort != null)
            {
                if (serialPort.IsOpen)
                {
                    try
                    {
                        mut2.WaitOne();
                        if (radioButton1.Checked)
                        {
                            serialPort.Write("ATDATA\n");
                            raw_data_st = serialPort.ReadLine().Split(',');
                            mut2.ReleaseMutex();
                            raw_data_st[raw_data_st.Length - 1] = raw_data_st[raw_data_st.Length - 1].Split(' ')[1];
                            raw_data = Array.ConvertAll(raw_data_st, int.Parse);
                            int[][] data_xy = new int[18][];
                            for (int i = 0; i < 18; i++)
                            {
                                data_xy[i] = new int[] { wavelenghts[i], raw_data[i] };
                            }
                            data_xy = data_xy.OrderBy(entry => entry[0]).ToArray();
                            e.Result = data_xy;
                        }
                        else if (radioButton2.Checked)
                        {
                            serialPort.Write("ATCDATA\n");
                            raw_data_st = serialPort.ReadLine().Split(',');
                            mut2.ReleaseMutex();
                            raw_data_st[raw_data_st.Length - 1] = raw_data_st[raw_data_st.Length - 1].Split(' ')[1];
                            calibrated_data = Array.ConvertAll(raw_data_st, float.Parse);
                            float[][] data_xy = new float[18][];
                            for (int i = 0; i < 18; i++)
                            {
                                data_xy[i] = new float[] { wavelenghts[i], calibrated_data[i] };
                            }
                            data_xy = data_xy.OrderBy(entry => entry[0]).ToArray();
                            e.Result = data_xy;
                        }
                            
                    }catch(Exception ex)
                    {
                        mut2.ReleaseMutex();
                        if (ex is TimeoutException || ex is InvalidOperationException)
                        {
                            serial_close();
                            connect_info_label.Text = "COM not respond, try change COM or reconnect.";
                        }
                    }
                }
            }
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (serialPort != null && e.Error == null && e.Cancelled == false)
            {
                try {
                    if (Object.ReferenceEquals(e.Result.GetType(),typeof(int[][])))
                    {
                        int[][] data_xy = (int[][])e.Result;

                        series1.Points.Clear();
                        series2.Points.Clear();
                        int[] rgb;
                        Boolean flag_join = true;
                        for (int i = 0; i < 18; i++)
                        {
                            if (data_xy[i][0] < 740)
                            {
                                series1.Points.AddXY(data_xy[i][0], data_xy[i][1]);
                                rgb = waveLengthToRGB(data_xy[i][0]);
                                series1.Points[i].Color = Color.FromArgb(rgb[0], rgb[1], rgb[2]);
                                series1.Points[i].Label = (data_xy[i][0] + "|" + data_xy[i][1]).ToString();
                            }
                            else
                            {
                                if (flag_join)
                                {
                                    series2.Points.AddXY(data_xy[i - 1][0], data_xy[i - 1][1]);
                                    flag_join = false;
                                }
                                series2.Points.AddXY(data_xy[i][0], data_xy[i][1]);
                                series2.Points[i - 12].Label = (data_xy[i][0] + "|" + data_xy[i][1]).ToString();
                            }

                        }


                        int max_val = Int32.MinValue;
                        int min_val = Int32.MaxValue;
                        int max_pos = 0, min_pos = 0;
                        for (int i = 0; i < 18; i++)
                        {
                            if (data_xy[i][1] > max_val)
                            {
                                max_val = data_xy[i][1];
                                max_pos = i;
                            }
                            if (data_xy[i][1] < min_val)
                            {
                                min_val = data_xy[i][1];
                                min_pos = i;
                            }
                        }

                        label10.Text = "Max X : " + data_xy[max_pos][0] + " ,Max Y : " + max_val;
                        label11.Text = "Min X : " + data_xy[min_pos][0] + " ,Min Y : " + min_val;
                    }else if (Object.ReferenceEquals(e.Result.GetType(), typeof(float[][])))
                    {
                        float[][] data_xy = (float[][])e.Result;

                        series1.Points.Clear();
                        series2.Points.Clear();
                        int[] rgb;
                        Boolean flag_join = true;
                        for (int i = 0; i < 18; i++)
                        {
                            if (data_xy[i][0] < 740)
                            {
                                series1.Points.AddXY(data_xy[i][0], data_xy[i][1]);
                                rgb = waveLengthToRGB(data_xy[i][0]);
                                series1.Points[i].Color = Color.FromArgb(rgb[0], rgb[1], rgb[2]);
                                series1.Points[i].Label = (data_xy[i][0] + "|" + data_xy[i][1]).ToString();
                            }
                            else
                            {
                                if (flag_join)
                                {
                                    series2.Points.AddXY(data_xy[i - 1][0], data_xy[i - 1][1]);
                                    flag_join = false;
                                }
                                series2.Points.AddXY(data_xy[i][0], data_xy[i][1]);
                                series2.Points[i - 12].Label = (data_xy[i][0] + "|" + data_xy[i][1]).ToString();
                            }

                        }


                        float max_val = Int32.MinValue;
                        float min_val = Int32.MaxValue;
                        int max_pos = 0, min_pos = 0;
                        for (int i = 0; i < 18; i++)
                        {
                            if (data_xy[i][1] > max_val)
                            {
                                max_val = data_xy[i][1];
                                max_pos = i;
                            }
                            if (data_xy[i][1] < min_val)
                            {
                                min_val = data_xy[i][1];
                                min_pos = i;
                            }
                        }

                        label10.Text = "Max X : " + data_xy[max_pos][0] + " , Max Y : " + max_val;
                        label11.Text = "Min X : " + data_xy[min_pos][0] + " , Min Y : " + min_val;
                    }
                    Thread.Sleep(10); //Sleep 10msec
                    if (sample_continue && serialPort != null)
                        backgroundWorker1.RunWorkerAsync();
                }catch(Exception ex)
                {
                    serial_close();
                    connect_info_label.Text = "COM not respond, try change COM or reconnect.";
                }

            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Stop the background worker thread (if running) to avoid race hazard.
            if (backgroundWorker1.IsBusy)
            {
                sample_continue = false;
                backgroundWorker1.CancelAsync();

                // Wait for the background worker thread to actually finish.
                while (backgroundWorker1.IsBusy)
                {
                    Application.DoEvents();
                    Thread.Sleep(100);
                }
            }
        }
        //**************************************************************************************************************************
        //---------------------------------CHART FUNCTIONS--------------------------------------------------------------------------
        private void Chart_MouseWheel(object sender, MouseEventArgs e)
        {
            var chart = (Chart)sender;
            var xAxis = chart.ChartAreas[0].AxisX;
            var yAxis = chart.ChartAreas[0].AxisY;

            try
            {
                if (e.Delta < 0) // Scrolled down.
                {
                    xAxis.ScaleView.ZoomReset();
                    yAxis.ScaleView.ZoomReset();
                }
                else if (e.Delta > 0) // Scrolled up.
                {
                    var xMin = xAxis.ScaleView.ViewMinimum;
                    var xMax = xAxis.ScaleView.ViewMaximum;
                    var yMin = yAxis.ScaleView.ViewMinimum;
                    var yMax = yAxis.ScaleView.ViewMaximum;

                    var posXStart = xAxis.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 4;
                    var posXFinish = xAxis.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 4;
                    var posYStart = yAxis.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 4;
                    var posYFinish = yAxis.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 4;

                    xAxis.ScaleView.Zoom(posXStart, posXFinish);
                    yAxis.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch { }
        }
        
        private void Chart_MouseMove(object sender, MouseEventArgs e)
        {
            if (tt == null) tt = new ToolTip();

            if (InnerPlotPositionClientRectangle(chart, chartArea1).Contains(e.Location))
            {

                Axis ax = chartArea1.AxisX;
                Axis ay = chartArea1.AxisY;
                double x = ax.PixelPositionToValue(e.X);
                double y = ay.PixelPositionToValue(e.Y);
                if (e.Location != tl)
                    tt.SetToolTip(chart, string.Format("X={0:0.00} , Y={1:0.00}", x, y));
                tl = e.Location;
            }
            else tt.Hide(chart);
        }
        RectangleF ChartAreaClientRectangle(Chart chart, ChartArea CA)
        {
            RectangleF CAR = CA.Position.ToRectangleF();
            float pw = chart.ClientSize.Width / 100f;
            float ph = chart.ClientSize.Height / 100f;
            return new RectangleF(pw * CAR.X, ph * CAR.Y, pw * CAR.Width, ph * CAR.Height);
        }

        RectangleF InnerPlotPositionClientRectangle(Chart chart, ChartArea CA)
        {
            RectangleF IPP = CA.InnerPlotPosition.ToRectangleF();
            RectangleF CArp = ChartAreaClientRectangle(chart, CA);

            float pw = CArp.Width / 100f;
            float ph = CArp.Height / 100f;

            return new RectangleF(CArp.X + pw * IPP.X, CArp.Y + ph * IPP.Y,
                                    pw * IPP.Width, ph * IPP.Height);
        }
        //**************************************************************************************************************************
        //---------------------------------OTHER FUNCTIONS--------------------------------------------------------------------------

        private int[] waveLengthToRGB(double Wavelength)
        {
            double factor;
            double Red, Green, Blue;

            if ((Wavelength >= 380) && (Wavelength < 440))
            {
                Red = -(Wavelength - 440) / (440 - 380);
                Green = 0.0;
                Blue = 1.0;
            }
            else if ((Wavelength >= 440) && (Wavelength < 490))
            {
                Red = 0.0;
                Green = (Wavelength - 440) / (490 - 440);
                Blue = 1.0;
            }
            else if ((Wavelength >= 490) && (Wavelength < 510))
            {
                Red = 0.0;
                Green = 1.0;
                Blue = -(Wavelength - 510) / (510 - 490);
            }
            else if ((Wavelength >= 510) && (Wavelength < 580))
            {
                Red = (Wavelength - 510) / (580 - 510);
                Green = 1.0;
                Blue = 0.0;
            }
            else if ((Wavelength >= 580) && (Wavelength < 645))
            {
                Red = 1.0;
                Green = -(Wavelength - 645) / (645 - 580);
                Blue = 0.0;
            }
            else if ((Wavelength >= 645) && (Wavelength < 781))
            {
                Red = 1.0;
                Green = 0.0;
                Blue = 0.0;
            }
            else
            {
                Red = 0.0;
                Green = 0.0;
                Blue = 0.0;
            }

            // Let the intensity fall off near the vision limits

            if ((Wavelength >= 380) && (Wavelength < 420))
                factor = 0.3 + 0.7 * (Wavelength - 380) / (420 - 380);
            else if ((Wavelength >= 420) && (Wavelength < 701))
                factor = 1.0;
            else if ((Wavelength >= 701) && (Wavelength < 781))
                factor = 0.3 + 0.7 * (780 - Wavelength) / (780 - 700);
            else
                factor = 0.0;
            
            int[] rgb = new int[3];

            // Don't want 0^x = 1 for x <> 0
            rgb[0] = Red == 0.0 ? 0 : (int)Math.Round(IntensityMax * Math.Pow(Red * factor, Gamma));
            rgb[1] = Green == 0.0 ? 0 : (int)Math.Round(IntensityMax * Math.Pow(Green * factor, Gamma));
            rgb[2] = Blue == 0.0 ? 0 : (int)Math.Round(IntensityMax * Math.Pow(Blue * factor, Gamma));

            return rgb;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = DateTime.Now.ToString("dd_MM_yyyy_hhmmss");
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ChartImageFormat cif = ChartImageFormat.Png;
                switch (saveFileDialog1.FilterIndex)
                {
                    case 1:
                        cif = ChartImageFormat.Png;
                        break;
                    case 2:
                        cif = ChartImageFormat.Jpeg;
                        break;
                    case 3:
                        cif = ChartImageFormat.Bmp;
                        break;
                    case 4:
                        cif = ChartImageFormat.Gif;
                        break;
                    case 5:
                        cif = ChartImageFormat.Tiff;
                        break;
                }
                chart.SaveImage(saveFileDialog1.FileName, cif);
            }
        }
    }
}