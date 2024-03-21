using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Net.NetworkInformation;
using EasyModbus;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Linq;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Table;
using OfficeOpenXml.Export.ToDataTable;
using System.ComponentModel;
using System.Reflection;

namespace WindowsFormsApp2
{
    public partial class Form2 : Form
    {
        private CancellationTokenSource cancellationTokenSource2;
        private CancellationTokenSource cancellationTokenSource3;
        private CancellationTokenSource cancellationTokenSource4;
        private CancellationTokenSource cancellationTokenSource5;
        private CancellationTokenSource cancellationTokenSource6;
        private CancellationTokenSource cancellationTokenSource7;
        private CancellationTokenSource cancellationTokenSource8;
        string V, G, H,K,L,U;
        DateTime Day;
        DateTime targetTime;
        DateTime Day2;
        DateTime targetTime2;
        DateTime Day3;
        DateTime targetTime3;
        DateTime Day4;
        DateTime targetTime4;
        DateTime Day5;
        DateTime targetTime5;
        DateTime Day6;
        DateTime targetTime6;
        List<string> fileContent2 = new List<string>();
        List<string> fileContent = new List<string>();
        List<string> fileContent3 = new List<string>();
        List<string> fileContent4 = new List<string>();
        List<string> fileContent5 = new List<string>();
        string fileName;
        string fileName2;
        string fileName3;
        bool Unconnect = false;
        bool Unconnect1 = false;
        bool Unconnect2 = false;
        bool write1 = false;
        bool write2 = false;
        bool write3 = false;

        int disconnect_count = 0;
        int disconnect_count2 = 0;
        int disconnect_count3 = 0;
        string IP1;
        string IP2;
        string IP3;
        short port;
        int TimeoutS=60000;


        public Form2()
        {
            InitializeComponent();
        }
        ModbusClient plc = new ModbusClient();
        ModbusClient plc1 = new ModbusClient();
        ModbusClient plc2 = new ModbusClient();

        private void Form2_Load(object sender, EventArgs e)
        {

            Load_time();
            button3.Enabled = false;
            button9.Enabled = false;
            button8.Enabled = false;
            button4.Enabled = false;
            btmConnect.Enabled = false;
            comboBox1.Items.Add("Vierwer1");
            comboBox1.Items.Add("Vierwer2");
            comboBox1.Items.Add("Vierwer3");
            comboBox1.SelectedIndex = 0;

            V = string.Join(",", "Datatime", "V_a", "V_b", "V_c", "I_a", "I_b", "I_c", "kW_a", "kW_b", "kW_c", "kVAR_a", "kVAR_b", "kVAR_c", "kVA_a", "kVA_b", "kVA_c", "PF_a", "PF_b", "PF_c", "kWh_a", "kWh_b", "kWh_c", "kVARh_a", "kVARh_b", "kVARh_c", "kVAh_a", "kVAh_b", "kVAh_c", "V_avg", "I_avg", "kW_tot", "kVAR_tot", "kVA_tot", "PF_tot", "kWh_tot", "kVARh_tot", "kVAh_tot", "Freq_a", "Freq_b", "Freq_c", "Freq_max");
            // V = string.Join(",", "Datatime", "name", "V", "I", "kW", "kVAR", "kVA", "PF", "kWh", "kVARh", "kVAh", "Freq", "V_avg", "I_avg", "kW_tot", "kVAR_tot", "kVA_tot", "PF_tot", "kWh_tot", "kVARh_tot", "kVAh_tot", "Freq_max");
            H = string.Join(",", "Datatime", "V_AB", "I_A", "V_BC", "I_B", "V_CA", "I_C");
            //G = string.Join(",", "Datatime", "KW_A", "KWH_A", "KW_B", "KWH_B", "KW_C", "KWH_C");

            fileContent2.Add(V);
            fileContent3.Add(H);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;




        }

        private void btmConnect_Click(object sender, EventArgs e)
        {
            try
            {
                TimeoutS = Convert.ToInt32(txtTimeout.Text);
            }
            catch (Exception)
            {

            }
            switch (comboBox1.SelectedIndex)
            {



                case 0:
                    {
                        try
                        {
                            int timeout = Convert.ToInt32(txtTimeout.Text.Trim());
                            bool isPing = PingHost(txtIP.Text.Trim(), timeout);
                            if (isPing == false)
                                MessageBox.Show("無法連接", "錯誤", MessageBoxButtons.OK);
                                
                            else
                            {
                                if (Unconnect == false)
                                {
                                    Unconnect = true;
                                    plc?.Disconnect();
                                    IP1 = txtIP.Text.Trim();
                                    port = Convert.ToInt16(txtPort.Text.Trim());
                                    plc.IPAddress = txtIP.Text.Trim();
                                    plc.Port = Convert.ToInt16(txtPort.Text.Trim());

                                    plc.Connect();
                                   
                                    setStatusLabel();
                               



                                    if (dataGridView1.RowCount <= 0)


                                    {
                                        dataGridView1.Rows.Clear();
                                        string line = string.Empty;
                                        int rr = 0;
                                        using (StreamReader sr = new StreamReader(textBox1.Text))
                                        {
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                string[] ar = line.Split(',');
                                                if (rr > 0) dataGridView1.Rows.Add(ar[0], ar[1], ar[2], String.Empty);
                                                rr++;
                                            }
                                        }
                                    }
                                    btmConnect.Text = "中斷";
                                    if (plc.Connected == true)
                                        MessageBox.Show("PLC連線成功。", "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    label1.Text = txtIP.Text.Trim();
                                    ReadPLC();
                                }

                                else
                                {
                                    if (plc.Connected == true)
                                    {


                                        MessageBox.Show("連線失敗次數" + disconnect_count, "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        cancellationTokenSource2.Cancel();
                                        btmConnect.Text = "開始連線";
                                        Unconnect = false;
                                        plc?.Disconnect();
                                        setStatusLabel();
                                    }
                                }


                            }
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {

                        }
                    }


                    break;
                case 1:
                    {
                        try
                        {
                            int timeout1 = Convert.ToInt32(txtTimeout.Text.Trim());
                            bool isPing1 = PingHost(txtIP.Text.Trim(), timeout1);
                            if (isPing1 == false)
                                MessageBox.Show("無法連接", "錯誤", MessageBoxButtons.OK);
                            else
                            {
                                if (Unconnect1 == false)
                                {
                                    Unconnect1 = true;
                                    plc1?.Disconnect();
                                    IP2 = txtIP.Text.Trim();
                                    port = Convert.ToInt16(txtPort.Text.Trim());
                                    plc1.IPAddress = txtIP.Text.Trim();
                                    plc1.Port = Convert.ToInt16(txtPort.Text.Trim());

                                    plc1.Connect();
                                    setStatusLabel();



                                    if (dataGridView2.RowCount <= 0)

                                    {
                                        dataGridView2.Rows.Clear();
                                        string line = string.Empty;
                                        int rr = 0;
                                        using (StreamReader sr = new StreamReader(textBox1.Text))
                                        {
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                string[] ar = line.Split(',');
                                                if (rr > 0) dataGridView2.Rows.Add(ar[0], ar[1], ar[2], String.Empty);
                                                rr++;
                                            }
                                        }









                                    }
                                    btmConnect.Text = "中斷";
                                    if (plc1.Connected == true)
                                        MessageBox.Show("PLC連線成功。", "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    ReadPLC2();

                                    label2.Text = txtIP.Text.Trim();

                                }

                                else
                                {
                                    if (plc1.Connected == true)
                                    {
                                        MessageBox.Show("連線失敗次數" + disconnect_count2, "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        cancellationTokenSource4.Cancel();
                                        btmConnect.Text = "開始連線";
                                        Unconnect1 = false;
                                        plc1?.Disconnect();
                                        setStatusLabel();

                                    }




                                }

                            }

                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {

                        }
                    }
                    break;

                case 2:
                    {
                        try
                        {
                            int timeout2 = Convert.ToInt32(txtTimeout.Text.Trim());
                            bool isPing2 = PingHost(txtIP.Text.Trim(), timeout2);
                            if (isPing2 == false)
                                MessageBox.Show("無法連接", "錯誤", MessageBoxButtons.OK);
                            else
                            {
                                if (Unconnect2 == false)
                                {
                                    Unconnect2 = true;
                                    plc2?.Disconnect();
                                    IP3 = txtIP.Text.Trim();
                                    port = Convert.ToInt16(txtPort.Text.Trim());
                                    plc2.IPAddress = txtIP.Text.Trim();
                                    plc2.Port = Convert.ToInt16(txtPort.Text.Trim());

                                    plc2.Connect();
                                    setStatusLabel();



                                    if (dataGridView3.RowCount <= 0)

                                    {
                                        dataGridView3.Rows.Clear();
                                        string line = string.Empty;
                                        int rr = 0;
                                        using (StreamReader sr = new StreamReader(textBox1.Text))
                                        {
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                string[] ar = line.Split(',');
                                                if (rr > 0) dataGridView3.Rows.Add(ar[0], ar[1], ar[2], String.Empty);
                                                rr++;
                                            }
                                        }









                                    }
                                    btmConnect.Text = "中斷";
                                    if (plc2.Connected == true)
                                        MessageBox.Show("PLC連線成功。", "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    ReadPLC3();
                                    label4.Text = txtIP.Text.Trim();


                                }

                                else
                                {
                                    if (plc2.Connected == true)
                                    {
                                        MessageBox.Show("連線失敗次數" + disconnect_count3, "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        cancellationTokenSource5.Cancel();
                                        btmConnect.Text = "開始連線";
                                        Unconnect2 = false;
                                        plc2?.Disconnect();
                                        setStatusLabel();

                                    }



                                }
                            }

                        }
                        catch (Exception ex)
                        {


                        }

                    }
                    break;

            }

        }
         private void TryReconnectPLC()
        {
            try
            {
                plc.Disconnect();
                // 重新連線 PLC 的代碼
                plc.Connect();
                // 更新連線狀態顯示
                // 重新連線後，可以進行其他必要的初始化操作
                // 例如重新讀取配置，重新啟動相應的執行緒等等
            }
            catch (Exception ex)
            {
                // 重新連線失敗，可以記錄錯誤或進一步處理
                WriteToLogFile($"Reconnection failed: {ex.Message}");
            }
        }
        public void WriteToLogFile(string message)
        {
            string logFileName = "Logs\\ErrorFile.txt"; // 指定日志文件的路径
            string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logFileName);
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                // 写入信息到文件
                writer.WriteLine(DateTime.Now + ": " + message);
            }
        }
        private void setStatusLabel()
        {
            switch (comboBox1.SelectedIndex)
            {



                case 0:
                    {
                        bool connect = (plc == null) ? false : plc.Connected;

                        Lbstatus.Text = string.Format("連線狀態： {0}", connect);
                        break;
                    }
                case 1:
                    {
                        bool connect = (plc1 == null) ? false : plc1.Connected;

                        Lbstatus.Text = string.Format("連線狀態： {0}", connect);

                        break;
                    }
                case 2:
                    {
                        bool connect = (plc2 == null) ? false : plc2.Connected;

                        Lbstatus.Text = string.Format("連線狀態： {0}", connect);

                        break;
                    }
            }
        }

        public static bool PingHost(string _ip, int Timeout) //判斷IP是否還可以連線
        {

            bool pingable = false;
            Ping pinger = null;

            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(_ip, timeout: Timeout);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }

            return pingable;

        }



        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                btmConnect.Enabled = true;
            }

        }

        private void ReadPLC()
        {
            int y = 0;

            if (int.TryParse(txtTimeout.Text, out y))
            {
                cancellationTokenSource2 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource2.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                break;
                            }

                            try
                            {
                                int timeout = Convert.ToInt32(txtTimeout.Text.Trim());
                                bool isPing = PingHost(txtIP.Text.Trim(), timeout);
                                if (isPing == false)
                                {
                                    TryReconnectPLC();
                                    disconnect_count += 1;
                                    //Lbstatus.Text = string.Format("連線狀態： {0}", "連線中斷");

                                    //IP1 = txtIP.Text.Trim();
                                    //port = Convert.ToInt16(txtPort.Text.Trim());
                                    //bool isPing2 = PingHost(txtIP.Text.Trim(), timeout);
                                    //if (isPing2 == false)
                                    //{
                                    //    plc?.Disconnect();
                                    //    plc.Connect();
                                    //}
                                   
                                    //Lbstatus.Text = string.Format("連線狀態： {0}", "再次連線");
                                }

                                else
                                {
                                    DateTime currentTime = DateTime.Now;
                                    string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");

                                    if (!plc.Connected)
                                    {
                                        // 如果 PLC 斷線，重新連線
                                        TryReconnectPLC();
                                    }
                                    // 建立一個 List 來存儲所有的 Modbus 通訊結果
                                    List<int[]> modbusResults = new List<int[]>();

                                    // 第一次迴圈：進行 Modbus 通訊的讀取操作，並將結果存入 modbusResults
                                    foreach (DataGridViewRow Row in dataGridView1.Rows)
                                    {
                                        int address = Convert.ToInt32(Row.Cells[1].Value);
                                        int[] x = plc.ReadHoldingRegisters(address, 2);
                                        modbusResults.Add(x);
                                    }
                                    if (!plc.Connected)
                                    {
                                        // 如果 PLC 斷線，重新連線
                                        TryReconnectPLC();
                                    }
                                    // 第二次迴圈：將讀取到的 Modbus 資料分配給每一列的儲存格
                                    int rowIndex = 0;
                                    foreach (DataGridViewRow Row in dataGridView1.Rows)
                                    {
                                        int[] x = modbusResults[rowIndex];

                                        // 假設要將結果存入儲存格的第四個欄位（即 Cells[3]）
                                        Row.Cells[3].Value = ModbusClient.ConvertRegistersToFloat(x).ToString();

                                        rowIndex++;
                                    }

                                    // 更新顯示
                                    dataGridView1.Invalidate();







                                    Thread.Sleep(y);

                                }
                            }
                            catch (Exception ex)
                            {
                                try { 
                                string xx = ex.Message;
                                WriteToLogFile(xx);
                                }
                                catch { ex.ToString(); }
                            }
                            finally
                            {


                            }
                            if (cancellationTokenSource2.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                break;
                            }
                        }
                    }, cancellationTokenSource2.Token);
                });

            }
            else
            {
                // 轉換失敗的處理邏輯
                // 在這裡可以顯示錯誤訊息，或者使用預設值，或者執行其他適當的處理
                MessageBox.Show("無效的計時器間隔設定。請輸入有效的數字。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



            //foreach (DataGridViewRow Row in dataGridView1.Rows)
            //{



            //    for (int i = 0; i < Row.Cells.Count; i++)
            //    {

            //        int address = Convert.ToInt32(Row.Cells[1].Value);
            //        int[] x = plc.ReadHoldingRegisters(address, 2);


            //        Row.Cells[3].Value = ModbusClient.ConvertRegistersToFloat(x).ToString() ;
            //        dataGridView1.Invalidate();
            //    }


            //}




        }
        private void ReadPLC2()
        {
            int y = 0;

            if (int.TryParse(txtTimeout.Text, out y))
            {
                cancellationTokenSource4 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource4.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                break;
                            }

                            try
                            {
                                int timeout1 = Convert.ToInt32(txtTimeout.Text.Trim());
                                bool isPing = PingHost(txtIP.Text.Trim(), timeout1);
                                if (isPing == false)
                                {
                                    TryReconnectPLC();
                                    disconnect_count2 += 1;
                                    //Lbstatus.Text = string.Format("連線狀態： {0}", "連線中斷");
                                   
                                    //Lbstatus.Text = string.Format("連線狀態： {0}", "再次連線");
                                }

                                else
                                {
                                    DateTime currentTime = DateTime.Now;
                                    string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");

                                    // 建立一個 List 來存儲所有的 Modbus 通訊結果
                                    List<int[]> modbusResults = new List<int[]>();

                                    // 第一次迴圈：進行 Modbus 通訊的讀取操作，並將結果存入 modbusResults
                                    foreach (DataGridViewRow Row in dataGridView2.Rows)
                                    {
                                        int address = Convert.ToInt32(Row.Cells[1].Value);
                                        int[] x = plc1.ReadHoldingRegisters(address, 2);
                                        modbusResults.Add(x);
                                    }

                                    // 第二次迴圈：將讀取到的 Modbus 資料分配給每一列的儲存格
                                    int rowIndex = 0;
                                    foreach (DataGridViewRow Row in dataGridView2.Rows)
                                    {
                                        int[] x = modbusResults[rowIndex];

                                        // 假設要將結果存入儲存格的第四個欄位（即 Cells[3]）
                                        Row.Cells[3].Value = ModbusClient.ConvertRegistersToFloat(x).ToString();

                                        rowIndex++;
                                    }

                                    // 更新顯示
                                    dataGridView2.Invalidate();







                                    Thread.Sleep(y);

                                }
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    string xx = ex.Message;
                                    WriteToLogFile(xx);
                                }
                                catch { ex.ToString(); }
                            }
                            finally
                            {


                            }
                            if (cancellationTokenSource4.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                break;
                            }
                        }
                    }, cancellationTokenSource4.Token);
                });

            }
            else
            {
                // 轉換失敗的處理邏輯
                // 在這裡可以顯示錯誤訊息，或者使用預設值，或者執行其他適當的處理
                MessageBox.Show("無效的計時器間隔設定。請輸入有效的數字。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ReadPLC3()
        {
            int y = 0;

            if (int.TryParse(txtTimeout.Text, out y))
            {
                cancellationTokenSource5 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource5.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                break;
                            }

                            try
                            {
                                int timeout2 = Convert.ToInt32(txtTimeout.Text.Trim());
                                bool isPing = PingHost(txtIP.Text.Trim(), timeout2);
                                if (isPing == false)
                                {
                                    TryReconnectPLC();
                                    disconnect_count3 += 1;
                                    //Lbstatus.Text = string.Format("連線狀態： {0}", "連線中斷");
                                    
                                    //Lbstatus.Text = string.Format("連線狀態： {0}", "再次連線");
                                }

                                else
                                {
                                    DateTime currentTime = DateTime.Now;
                                    string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");

                                    // 建立一個 List 來存儲所有的 Modbus 通訊結果
                                    List<int[]> modbusResults = new List<int[]>();

                                    // 第一次迴圈：進行 Modbus 通訊的讀取操作，並將結果存入 modbusResults
                                    foreach (DataGridViewRow Row in dataGridView3.Rows)
                                    {
                                        int address = Convert.ToInt32(Row.Cells[1].Value);
                                        int[] x = plc2.ReadHoldingRegisters(address, 2);
                                        modbusResults.Add(x);
                                    }

                                    // 第二次迴圈：將讀取到的 Modbus 資料分配給每一列的儲存格
                                    int rowIndex = 0;
                                    foreach (DataGridViewRow Row in dataGridView3.Rows)
                                    {
                                        int[] x = modbusResults[rowIndex];

                                        // 假設要將結果存入儲存格的第四個欄位（即 Cells[3]）
                                        Row.Cells[3].Value = ModbusClient.ConvertRegistersToFloat(x).ToString();

                                        rowIndex++;
                                    }

                                    // 更新顯示
                                    dataGridView3.Invalidate();







                                    Thread.Sleep(y);

                                }
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    string xx = ex.Message;
                                    WriteToLogFile(xx);
                                }
                                catch { ex.ToString(); }
                            }
                            finally
                            {


                            }
                            if (cancellationTokenSource5.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                break;
                            }
                        }
                    }, cancellationTokenSource5.Token);
                });

            }
            else
            {
                // 轉換失敗的處理邏輯
                // 在這裡可以顯示錯誤訊息，或者使用預設值，或者執行其他適當的處理
                MessageBox.Show("無效的計時器間隔設定。請輸入有效的數字。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //int x = 0;
        //x = Convert.ToInt16(TimerInterval.Text);

        //if (int.TryParse(TimerInterval.Text, out x))
        //{
        //    timer2.Interval = x;
        //    SaveInformation();

        //}
        //else
        //{
        //    // 轉換失敗的處理邏輯
        //    // 在這裡可以顯示錯誤訊息，或者使用預設值，或者執行其他適當的處理
        //    MessageBox.Show("無效的計時器間隔設定。請輸入有效的數字。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //}


        private void SaveInformation()
        {
            try
            {

                int s = 60000;
                string[] ar = new string[40];
                string a, b, c = "";
               
                cancellationTokenSource3 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {

                            // 檢查取消請求
                            if (cancellationTokenSource3.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource3.Token))
                                {
                                    //write1 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row2 in dataGridView1.Rows)
                            {

                                for (int i = 0; i < 40; i++)
                                {


                                    ar[i] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                                }

                            }
                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");




                            Day = DateTime.Now;
                            if (Day > targetTime)
                            {
                                
                                string f = "";
                                f = currentTime.ToString("yyyyMMddHHmmss") + ".csv";
                                fileName = f;
                                fileContent.Add(V);
                                Day = DateTime.Now;
                                targetTime = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天

                            }

                            // 設定預設的檔案副檔名及檔案類型篩選
                            //   V = string.Join(",", "Datatime", "V_a", "V_b", "V_c", "I_a", "I_b", "I_c", "kW_a", "kW_b", "kW_c", "kVAR_a", "kVAR_b", "kVAR_c", "kVA_a", "kVA_b", "kVA_c", "PF_a", "PF_b", "PF_c", "kWh_a", "kWh_b", "kWh_c", "kVARh_a", "kVARh_b", "kVARh_c", "kVAh_a", "kVAh_b", "kVAh_c", "V_avg", "I_avg", "kW_tot", "kVAR_tot", "kVA_tot", "PF_tot", "kWh_tot", "kVARh_tot", "kVAh_tot", "Freq_a", "Freq_b", "Freq_c", "Freq_max");
                            a = string.Join(",", formattedTime, ar[0], ar[9], ar[18], ar[1], ar[10], ar[19], ar[2], ar[11], ar[20], ar[3], ar[12], ar[21], ar[4], ar[13], ar[22], ar[5], ar[14], ar[23], ar[6], ar[15], ar[24], ar[7], ar[16], ar[25], ar[8], ar[17], ar[26], ar[27], ar[28], ar[29], ar[30], ar[31], ar[32], ar[33], ar[34], ar[35], ar[36], ar[37], ar[38], ar[39]);
                            fileContent.Add(a);

                            //using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\Work\test.xlsx"))) 
                            //{
                            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Data");//创建worksheet



                            //    package.Save();//保存excel



                            //var classList = new List<testClass>();

                            //foreach (var item in fileContent)
                            //{
                            //    var values = item.Split(',');

                            //    var classInstance = new testClass
                            //    {
                            //        Time = values[0],
                            //        name = values[1],
                            //        V_Value = values[2],
                            //        I_Value = values[3],
                            //        kW_Value = values[4],
                            //        kVAR_Value = values[5],
                            //        kVA_Value = values[6],
                            //        PF_Value = values[7],
                            //        kWh_Value = values[8],
                            //        kVARh_Value = values[9],
                            //        kVAh_Value = values[10],
                            //        Freq_Value = values[11],
                            //        V_avgValue = values[12],
                            //        I_avgValue = values[13],
                            //        kW_totValue = values[14],
                            //        kVAR_totValue = values[15],
                            //        kVA_totValue = values[16],
                            //        PF_totValue = values[17],
                            //        kWh_totValue = values[18],
                            //        kVARh_totValue = values[19],
                            //        kVAh_totValue = values[20],
                            //        Freq_maxValue = values[21]
                            //    };

                            //    classList.Add(classInstance);
                            //}
                            //ExportExcel(classList, fileName);

                         

                            //}
                            try
                            {
                                File.AppendAllLines(fileName, fileContent);
                                fileContent.Clear();
                                a = null;
                            }
                            catch (Exception ex)
                            {
                                string xx = ex.Message;
                              
                            }
                            // 檢查取消請求
                            if (cancellationTokenSource3.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource3.Token))
                                {
                                    //write1 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource3.Token);
                });
            }

            catch (Exception ex)
            {

            }
            finally
            {
                //fileContent.Clear();


            }

        }
        private void SaveInformation2()
        {
            try
            {
                int s = 60000;
                string[] ar2 = new string[40];
                string a, b, c = "";
               
                cancellationTokenSource6 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource6.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource6.Token))
                                {
                                    //write2 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row2 in dataGridView2.Rows)
                            {

                                for (int i = 0; i < 40; i++)
                                {


                                    ar2[i] = dataGridView2.Rows[i].Cells[3].Value.ToString();
                                }

                            }

                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");
                            Day2 = DateTime.Now;
                            if (Day2 > targetTime2)
                            {
                                string f = "";
                                f = currentTime.ToString("yyyyMMddHHmmss") + ".csv";
                                fileName2 = f;
                                fileContent4.Add(V);
                                Day = DateTime.Now;
                                targetTime2 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                            }
                            a = string.Join(",", formattedTime, ar2[0], ar2[9], ar2[18], ar2[1], ar2[10], ar2[19], ar2[2], ar2[11], ar2[20], ar2[3], ar2[12], ar2[21], ar2[4], ar2[13], ar2[22], ar2[5], ar2[14], ar2[23], ar2[6], ar2[15], ar2[24], ar2[7], ar2[16], ar2[25], ar2[8], ar2[17], ar2[26], ar2[27], ar2[28], ar2[29], ar2[30], ar2[31], ar2[32], ar2[33], ar2[34], ar2[35], ar2[36], ar2[37], ar2[38], ar2[39]);
                            //b = string.Join(",", formattedTime, "b", ar2[9], ar2[10], ar2[11], ar2[12], ar2[13], ar2[14], ar2[15], ar2[16], ar2[17], ar2[37], ar2[27], ar2[28], ar2[29], ar2[30], ar2[31], ar2[32], ar2[33], ar2[34], ar2[35], ar2[39]);
                            //c = string.Join(",", formattedTime, "c", ar2[18], ar2[19], ar2[20], ar2[21], ar2[22], ar2[23], ar2[24], ar2[25], ar2[26], ar2[38], ar2[27], ar2[28], ar2[29], ar2[30], ar2[31], ar2[32], ar2[33], ar2[34], ar2[35], ar2[39]);
                            fileContent4.Add(a);
                            //fileContent4.Add(b);
                            //fileContent4.Add(c);

                            //var classList = new List<testClass>();

                            //foreach (var item in fileContent4)
                            //{
                            //    var values = item.Split(',');

                            //    var classInstance = new testClass
                            //    {
                            //        Time = values[0],
                            //        name = values[1],
                            //        V_Value = values[2],
                            //        I_Value = values[3],
                            //        kW_Value = values[4],
                            //        kVAR_Value = values[5],
                            //        kVA_Value = values[6],
                            //        PF_Value = values[7],
                            //        kWh_Value = values[8],
                            //        kVARh_Value = values[9],
                            //        kVAh_Value = values[10],
                            //        Freq_Value = values[11],
                            //        V_avgValue = values[12],
                            //        I_avgValue = values[13],
                            //        kW_totValue = values[14],
                            //        kVAR_totValue = values[15],
                            //        kVA_totValue = values[16],
                            //        PF_totValue = values[17],
                            //        kWh_totValue = values[18],
                            //        kVARh_totValue = values[19],
                            //        kVAh_totValue = values[20],
                            //        Freq_maxValue = values[21]
                            //    };

                            //    classList.Add(classInstance);
                            //}
                            //ExportExcel(classList, fileName2);
                 
                            try
                            {
                                File.AppendAllLines(fileName2, fileContent4);
                                fileContent4.Clear();
                                a = null;
                            }
                            catch (Exception ex)
                            {
                                string xx = ex.Message;
                                
                            }

                            //b = null;
                            //c = null;
                            // 檢查取消請求
                            if (cancellationTokenSource6.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource6.Token))
                                {
                                    //write2 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource6.Token);
                });

            }
            catch (Exception ex)
            {

            }
            finally
            {
                //fileContent4.Clear();


            }

        }
        private void SaveInformation3()
        {

            try
            {
                int s = 60000;
                string[] ar3 = new string[40];
                string a, b, c = "";
               
                cancellationTokenSource7 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource7.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource7.Token))
                                {
                                    //write3= false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row3 in dataGridView3.Rows)
                            {

                                for (int i = 0; i < 40; i++)
                                {


                                    ar3[i] = dataGridView3.Rows[i].Cells[3].Value.ToString();




                                }

                            }
                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");
                            Day3 = DateTime.Now;
                            if (Day3 > targetTime3)
                            {
                                string f = "";
                                f = currentTime.ToString("yyyyMMddHHmmss") + ".csv";
                                fileName3 = f;
                                fileContent5.Add(V);
                                Day3 = DateTime.Now;
                                targetTime3 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                            }

                            a = string.Join(",", formattedTime, ar3[0], ar3[9], ar3[18], ar3[1], ar3[10], ar3[19], ar3[2], ar3[11], ar3[20], ar3[3], ar3[12], ar3[21], ar3[4], ar3[13], ar3[22], ar3[5], ar3[14], ar3[23], ar3[6], ar3[15], ar3[24], ar3[7], ar3[16], ar3[25], ar3[8], ar3[17], ar3[26], ar3[27], ar3[28], ar3[29], ar3[30], ar3[31], ar3[32], ar3[33], ar3[34], ar3[35], ar3[36], ar3[37], ar3[38], ar3[39]);
                            //b = string.Join(",", formattedTime, "b", ar3[9], ar3[10], ar3[11], ar3[12], ar3[13], ar3[14], ar3[15], ar3[16], ar3[17], ar3[37], ar3[27], ar3[28], ar3[29], ar3[30], ar3[31], ar3[32], ar3[33], ar3[34], ar3[35], ar3[39]);
                            //c = string.Join(",", formattedTime, "c", ar3[18], ar3[19], ar3[20], ar3[21], ar3[22], ar3[23], ar3[24], ar3[25], ar3[26], ar3[38], ar3[27], ar3[28], ar3[29], ar3[30], ar3[31], ar3[32], ar3[33], ar3[34], ar3[35], ar3[39]);
                            fileContent5.Add(a);
                            //fileContent5.Add(b);
                            //fileContent5.Add(c);

                            //if (File.Exists(fileName))
                            //{
                            //    var classList = new List<testClass>();

                            //    foreach (var item in fileContent5)
                            //    {
                            //        var values = item.Split(',');

                            //        var classInstance = new testClass
                            //        {
                            //            Time = values[0],
                            //            name = values[1],
                            //            V_Value = values[2],
                            //            I_Value = values[3],
                            //            kW_Value = values[4],
                            //            kVAR_Value = values[5],
                            //            kVA_Value = values[6],
                            //            PF_Value = values[7],
                            //            kWh_Value = values[8],
                            //            kVARh_Value = values[9],
                            //            kVAh_Value = values[10],
                            //            Freq_Value = values[11],
                            //            V_avgValue = values[12],
                            //            I_avgValue = values[13],
                            //            kW_totValue = values[14],
                            //            kVAR_totValue = values[15],
                            //            kVA_totValue = values[16],
                            //            PF_totValue = values[17],
                            //            kWh_totValue = values[18],
                            //            kVARh_totValue = values[19],
                            //            kVAh_totValue = values[20],
                            //            Freq_maxValue = values[21]
                            //        };

                            //        classList.Add(classInstance);
                            //    }
                            //    ExportExcel(classList, fileName3);
                            //}


                          


                            try
                            {
                                File.AppendAllLines(fileName3, fileContent5);
                                fileContent5.Clear();
                                a = null;
                            }
                            catch (Exception ex)
                            {
                                string xx = ex.Message;
                               
                            }


                            //b = null;
                            //c = null;
                            // 檢查取消請求
                            if (cancellationTokenSource7.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource7.Token))
                                {
                                    //write3 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource7.Token);
                });

            }
            catch (Exception ex)
            {
                string xx = ex.Message;
                
            }
            finally
            {
                //fileContent5.Clear();

            }



        }

        private void SaveInformation4()
        {
            try
            {

                int s = 60000;
                string[] ar = new string[17];
                string a = "";
                write1 = true;
                cancellationTokenSource3 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource3.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource3.Token))
                                {
                                    write1 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row2 in dataGridView1.Rows)
                            {

                                for (int i = 0; i < 17; i++)
                                {
                                    ar[i] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                                }

                            }
                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");


                            Day4 = DateTime.Now;
                            if (Day4 > targetTime4)
                            {
                                string f = "";
                                f = formattedTime.Replace("/", ".");
                                fileName = K + "\\" + label1.Text + "_" + f.Replace(":", " ") + ".csv";
                                fileContent.Add(H);
                                Day4 = DateTime.Now;
                                targetTime4 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                            }




                            // 設定預設的檔案副檔名及檔案類型篩選

                            a = string.Join(",", formattedTime, ar[0], ar[1], ar[2], ar[3], ar[4], ar[5], ar[6], ar[7], ar[8], ar[9], ar[10], ar[11], ar[12], ar[13], ar[14], ar[15], ar[16]);
                            fileContent.Add(a);




                            //}

                            File.AppendAllLines(fileName, fileContent);
                            fileContent.Clear();
                            a = null;

                            // 檢查取消請求
                            if (cancellationTokenSource3.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource3.Token))
                                {
                                    write1 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource3.Token);
                });
            }

            catch (Exception ex)
            {

            }
            finally
            {
                //fileContent.Clear();


            }

        }

        private void SaveInformation5()
        {
            try
            {

                int s = 60000;
                string[] ar = new string[17];
                string a = "";
                write2 = true;
                cancellationTokenSource6 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource6.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource6.Token))
                                {
                                    write2 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row2 in dataGridView2.Rows)
                            {

                                for (int i = 0; i < 17; i++)
                                {
                                    ar[i] = dataGridView2.Rows[i].Cells[3].Value.ToString();
                                }

                            }
                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");




                            Day5 = DateTime.Now;
                            if (Day5 > targetTime5)
                            {
                                string f = "";
                                f = formattedTime.Replace("/", ".");
                                fileName2 = L + "\\" + label2.Text + "_" + f.Replace(":", " ") + ".csv";
                                fileContent4.Add(H);
                                Day5 = DateTime.Now;
                                targetTime5 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                            }


                            // 設定預設的檔案副檔名及檔案類型篩選

                            a = string.Join(",", formattedTime, ar[0], ar[1], ar[2], ar[3], ar[4], ar[5], ar[6], ar[7], ar[8], ar[9], ar[10], ar[11], ar[12], ar[13], ar[14], ar[15], ar[16]);
                            fileContent4.Add(a);




                            //}

                            File.AppendAllLines(fileName2, fileContent4);
                            fileContent4.Clear();
                            a = null;

                            // 檢查取消請求
                            if (cancellationTokenSource6.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource6.Token))
                                {
                                    write2 = false;
                                    Thread.CurrentThread.Abort();
                                }
                           
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource6.Token);
                });
            }

            catch (Exception ex)
            {

            }
            finally
            {
                //fileContent.Clear();


            }

        }

        private void SaveInformation6()
        {
            try
            {

                int s = 60000;
                string[] ar = new string[17];
                string a = "";
                write3 = true;
                cancellationTokenSource7 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource7.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource7.Token))
                                {
                                    write3 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row2 in dataGridView3.Rows)
                            {

                                for (int i = 0; i < 17; i++)
                                {
                                    ar[i] = dataGridView3.Rows[i].Cells[3].Value.ToString();
                                }

                            }
                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");



                            Day6 = DateTime.Now;
                            if (Day6 > targetTime6)
                            {
                                //    Thread.Sleep(1000);
                                string f = "";
                                f = formattedTime.Replace("/", ".");
                                fileName3 = G + "\\" + label4.Text + "_" + f.Replace(":", " ") + ".csv";
                                fileContent5.Add(H);
                                Day6 = DateTime.Now;

                                targetTime6 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                            }


                            // 設定預設的檔案副檔名及檔案類型篩選

                            a = string.Join(",", formattedTime, ar[0], ar[1], ar[2], ar[3], ar[4], ar[5], ar[6], ar[7], ar[8], ar[9], ar[10], ar[11], ar[12], ar[13], ar[14], ar[15], ar[16]);
                            fileContent5.Add(a);



                    

                            //}

                            File.AppendAllLines(fileName3, fileContent5);
                            fileContent5.Clear();
                            a = null;

                            // 檢查取消請求
                            if (cancellationTokenSource7.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource7.Token))
                                {
                                    write3=false;
                                    Thread.CurrentThread.Abort();
                                }
                             
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource7.Token);
                });
            }

            catch (Exception ex)
            {

            }
            finally
            {
                //fileContent.Clear();


            }

        }
        private void SaveInformation7()
        {
            try
            {

                int s = 60000;
                string[] ar = new string[7];
                string a = "";
                write1 = true;
                cancellationTokenSource3 = new CancellationTokenSource();

                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {

                        while (true)
                        {
                            // 檢查取消請求
                            if (cancellationTokenSource3.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource3.Token))
                                {
                                    write1 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            foreach (DataGridViewRow Row2 in dataGridView1.Rows)
                            {

                                for (int i = 0; i < 6; i++)
                                {
                                    ar[i] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                                }

                            }
                            DateTime currentTime = DateTime.Now;
                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");


                            Day4 = DateTime.Now;
                            if (Day4 > targetTime4)
                            {
                                string f = "";
                                f = formattedTime.Replace("/", ".");
                                fileName = K + "\\" + label1.Text + "_" + f.Replace(":", " ") + ".csv";
                                fileContent.Add(H);
                                Day4 = DateTime.Now;
                                targetTime4 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                            }




                            // 設定預設的檔案副檔名及檔案類型篩選

                            a = string.Join(",", formattedTime, ar[0], ar[1], ar[2], ar[3], ar[4], ar[5], ar[6]);
                            fileContent.Add(a);
                            



                            //}

                            File.AppendAllLines(fileName, fileContent);
                            fileContent.Clear();
                            a = null;

                            // 檢查取消請求
                            if (cancellationTokenSource3.Token.IsCancellationRequested)
                            {
                                // 如果取消，則退出循環
                                if (Thread.CurrentThread.Equals(cancellationTokenSource3.Token))
                                {
                                    write1 = false;
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                            Thread.Sleep(TimeoutS);
                        }
                    }, cancellationTokenSource3.Token);
                });
            }

            catch (Exception ex)
            {

            }
            finally
            {
                //fileContent.Clear();


            }

        }

        private void Load_time()
        {
            try
            {
                cancellationTokenSource8 = new CancellationTokenSource();
                ThreadPool.QueueUserWorkItem(async state =>
                {

                    await Task.Run(() =>
                    {
                        while (true)
                        {
                            if (cancellationTokenSource8.Token.IsCancellationRequested)
                            {
                                if (Thread.CurrentThread.Equals(cancellationTokenSource8.Token))
                                {
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }
                            DateTime x = DateTime.Now;
                            string formattedTime = x.ToString("yyyy/MM/dd HH:mm:ss");
                            try
                            {
                                this.Invoke((MethodInvoker)delegate
                                {
                                    U = formattedTime;
                                    label12.Text = U;
                                });


                            }
                            catch
                            {

                            }

                            Thread.Sleep(1000);

                            if (cancellationTokenSource8.Token.IsCancellationRequested)
                            {
                                if (Thread.CurrentThread.Equals(cancellationTokenSource8.Token))
                                {
                                    Thread.CurrentThread.Abort();
                                }
                                break;
                            }

                        }

                    }, cancellationTokenSource8.Token);


                });
            }
            catch
            {
            }

        }
     


        private void button5_Click(object sender, EventArgs e)
        {
            SaveFile();



        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled=false;
            button4.Enabled=true;
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    {
                        write1 = true;
                        SaveInformation();
                    }
                    break;
                case 1:
                    {
                        write2 = true;
                        SaveInformation2();
                    }
                    break;
                case 2:
                    {
                        write3 = true;
                        SaveInformation3();
                    }
                    break;



            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            button3.Enabled = true;
            button4.Enabled = false;
            try
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        { cancellationTokenSource3.Cancel(); write1 = false; }
                        break;
                    case 1:
                        { cancellationTokenSource6.Cancel(); write2 = false; }
                        break;
                    case 2:
                        { cancellationTokenSource7.Cancel(); write3 = false; }
                        break;
                }
            }
            catch (Exception ex)
            {

            }


        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
          

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }


        public class testClass
        {
            [DisplayName("時間")]
            public string Time { set; get; }
            [DisplayName("名稱")]
            public string name { set; get; }
            [DisplayName("V")]
            public string V_Value { set; get; }
            [DisplayName("I")]
            public string I_Value { set; get; }
            [DisplayName("kW")]
            public string kW_Value { set; get; }
            [DisplayName("kVAR")]
            public string kVAR_Value { set; get; }
            [DisplayName("kVA")]
            public string kVA_Value { set; get; }
            [DisplayName("PF")]
            public string PF_Value { set; get; }
            [DisplayName("kWh")]
            public string kWh_Value { set; get; }
            [DisplayName("kVARh")]
            public string kVARh_Value { set; get; }
            [DisplayName("kVAh")]
            public string kVAh_Value { set; get; }
            [DisplayName("Freq")]
            public string Freq_Value { set; get; }
            [DisplayName("V_avg")]
            public string V_avgValue { set; get; }
            [DisplayName("I_avg")]
            public string I_avgValue { set; get; }
            [DisplayName("kW_tot")]
            public string kW_totValue { set; get; }
            [DisplayName("kVAR_tot")]
            public string kVAR_totValue { set; get; }
            [DisplayName("kVA_tot")]
            public string kVA_totValue { set; get; }
            [DisplayName("PF_tot")]
            public string PF_totValue { set; get; }
            [DisplayName("kWh_tot")]
            public string kWh_totValue { set; get; }
            [DisplayName("kVARh_tot")]
            public string kVARh_totValue { set; get; }
            [DisplayName("kVAh_tot")]
            public string kVAh_totValue { set; get; }
            [DisplayName("Freq_max")]
            public string Freq_maxValue { set; get; }


        }


        public FileInfo ExportExcel<T>(IEnumerable<T> data, string filename) where T : class
        {

            //var output = new MemoryStream();

            FileInfo output;

            if (File.Exists(filename))
            {
                // 如果文件已存在，打开现有文件
                output = new FileInfo(filename);
            }
            else
            {
                // 如果文件不存在，创建一个新文件
                output = new FileInfo(filename);
            }
            // 如果需要，创建工作表并设置表头
            //using (var excel = new ExcelPackage(output))
            //{
            //    var ws = excel.Workbook.Worksheets.Add("Sheet1");

            //    // 在此设置表头样式等
            //}

            //if (output.Exists)
            //{
            //    output.Delete();
            //    output = new FileInfo("C:\\ExportExcelTest-" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + ".xlsx");
            //}


            using (var excel = new ExcelPackage(output))
            {
                var ws = excel.Workbook.Worksheets.Add("Sheet1"); // 建立分頁

                // 用反射拿出有 DisplayName 的屬性
                var properties = typeof(T)
                    .GetProperties()
                    .Where(prop => prop.IsDefined(typeof(DisplayNameAttribute)));

                var rows = data.Count() + 1;   // 直：資料筆數（記得加標題列）
                var cols = properties.Count(); // 橫：類別中有別名的屬性數量

                if (rows > 0 && cols > 0)
                {
                    ws.Cells[1, 1].LoadFromCollection(data, true); // 寫入資料

                    // 儲存格格式
                    var colNumber = 1;
                    foreach (var prop in properties)
                    {
                        // 時間處理，如果沒指定儲存格格式會變成 通用格式，就會以 int＝時間戳 的方式顯示
                        if (prop.PropertyType.Equals(typeof(DateTime)) ||
                           prop.PropertyType.Equals(typeof(DateTime?)))
                        {
                            ws.Cells[2, colNumber, rows, colNumber].Style.Numberformat.Format = "mm-dd-yy hh:mm:ss";
                        }
                        colNumber += 1;
                    }

                    // 樣式準備
                    using (var range = ws.Cells[1, 1, rows, cols])
                    {
                        ws.Cells.Style.Font.Name = "新細明體";
                        ws.Cells.Style.Font.Size = 12;
                        ws.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 置中
                        ws.Cells.AutoFitColumns(); // 欄寬

                        // 框線
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        // 標題列
                        var title = ws.Cells[1, 1, 1, cols];
                        title.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; // 設定背景填色方法
                        title.Style.Fill.BackgroundColor.SetColor(color: System.Drawing.Color.AliceBlue);
                    }
                }
                else
                {

                }
                excel.Save(); // 儲存 Excel
            }
            //output.Position = 0; // 如果是使用 stream 的方式讓人下載，請記得將指標移回資料起始

            return output;
        }
        //SaveFileDialog saveFileDialog = new SaveFileDialog();


        //// 設定預設的檔案副檔名及檔案類型篩選
        //saveFileDialog.DefaultExt = "xlsx";
        //saveFileDialog.Filter = "xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
        //string vv,ss = "";
        //       DialogResult result = saveFileDialog.ShowDialog(); 
        //ss = saveFileDialog.FileName;


        //using (ExcelPackage package = new ExcelPackage(new FileInfo(@ss)))
        //{
        //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("1");//创建worksheet


        //    using (ExcelRange range = worksheet.Cells[1, 1, 5, 5])
        //    {
        //        range.Style.Font.Size = 12;
        //        range.Style.Font.Color.SetColor(color: System.Drawing.Color.AliceBlue);
        //        range.Style.Font.Name = "微软雅黑";//字体
        //        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; ;//水平居中
        //        range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;//垂直居中
        //        range.Style.WrapText = true;
        //        range.Style.ShrinkToFit = true;
        //        range.Worksheet.Row(1).CustomHeight = true;


        //        }

        //    package.Save();//保存excel
        //}
        private void button7_Click_1(object sender, EventArgs e)
        {
            //try
            //{
            //    // 如果要保存 CSV 文件
            //    if (Path.GetExtension(filename).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            //    {
            //        File.AppendAllLines(filename, fileContent);
            //        MessageBox.Show("CSV 文件已成功保存！", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //    // 如果要保存 XLSX 文件
            //    else if (Path.GetExtension(filename).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            //    {
            //        // 通过调用 ExportExcel 方法创建 Excel 文件
            //        FileInfo excelFile = ExportExcel(fileContent, filename);
            //        MessageBox.Show("XLSX 文件已成功保存！", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //    else
            //    {
            //        MessageBox.Show("不支持的文件类型！", "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("无法保存文件：" + ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFile2();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                txtIP.Text = plc.IPAddress;
                txtPort.Text = plc.Port.ToString();
                setStatusLabel();
                if (Unconnect == true)
                {
                    btmConnect.Text = "中斷";
                   
                    if (File.Exists(fileName))
                    {

                        label6.Text = txtIP.Text + "_" + Day.ToString("yyyy/MM/dd HH:mm:ss");
                        label7.Text = txtIP.Text + "_" + Day4.ToString("yyyy/MM/dd HH:mm:ss");
                        if (write1 == true)
                    {
                        button3.Enabled = false;
                        button4.Enabled = true;
                    }
                    else
                    {
                        button3.Enabled = true;
                        button4.Enabled = false;
                    }

                    }
                    else
                    {
                        button3.Enabled = false;
                        button4.Enabled = false;

                    }

                }
                else
                {
                    btmConnect.Text = "連線";
              
                }
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                txtIP.Text = plc1.IPAddress;
                txtPort.Text = plc1.Port.ToString();
                setStatusLabel();
                if (Unconnect1 == true)
                {
                    btmConnect.Text = "中斷";
                   
                    if (File.Exists(fileName2))
                    {

                        label6.Text = txtIP.Text + "_" + Day2.ToString("yyyy/MM/dd HH:mm:ss");
                        label7.Text = txtIP.Text + "_" + Day5.ToString("yyyy/MM/dd HH:mm:ss");
                        if (write2 == true)
                        {
                            button3.Enabled = false;
                            button4.Enabled = true;
                        }
                        else
                        {
                            button3.Enabled = true;
                            button4.Enabled = false;
                        }
                    }
                    else
                    {
                        button3.Enabled = false;
                        button4.Enabled = false;

                    }
                }
                else
                {
                    btmConnect.Text = "連線";
                }
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                txtIP.Text = plc2.IPAddress;
                txtPort.Text = plc2.Port.ToString();
                setStatusLabel();
                if (Unconnect2 == true)
                {
                 
                    btmConnect.Text = "中斷";
                    if (File.Exists(fileName3))
                    {

                        label6.Text = txtIP.Text + "_" + Day3.ToString("yyyy/MM/dd HH:mm:ss");
                        label7.Text = txtIP.Text + "_" + Day6.ToString("yyyy/MM/dd HH:mm:ss");
                         if (write3 == true)
                    {
                        button3.Enabled = false;
                        button4.Enabled = true;
                    }
                    else
                    {
                        button3.Enabled = true;
                        button4.Enabled = false;
                    } 
                    }
                    else
                    {
                        button3.Enabled = false;
                        button4.Enabled = false;

                    }
                }
                else
                {
                    btmConnect.Text = "連線";
               
                }
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            button9.Enabled = false;
            button8.Enabled = true;
            switch (comboBox1.SelectedIndex)
            {



                case 0:
                    {
                        SaveInformation4();
                        //SaveInformation7();
                        break;
                    }

                case 1:
                    {
                        SaveInformation5();
                        break;
                    }

                case 2:
                    {
                        SaveInformation6();
                        break;
                    }

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button8.Enabled = false;
            button9.Enabled = true;
            try
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        { cancellationTokenSource3.Cancel(); }
                        break;
                    case 1:
                        { cancellationTokenSource6.Cancel(); }
                        break;
                    case 2:
                        { cancellationTokenSource7.Cancel(); }
                        break;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            cancellationTokenSource8.Cancel();

            if (e.CloseReason == CloseReason.UserClosing)
            {
                // 觸發 MainForm 的 Show() 方法
                (Owner as Form3)?.Show();
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void SaveFile()
        {
            switch (comboBox1.SelectedIndex)
            {



                case 0:
                    {
                       
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        String a;
                        Day = DateTime.Now;
                        targetTime = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                        string formattedTime = Day.ToString("yyyy/MM/dd HH:mm:ss");
                        label6.Text = txtIP.Text + "_" + formattedTime;

                        // 設定預設的檔案副檔名及檔案類型篩選
                        a = formattedTime.Replace("/", ".");
                        saveFileDialog.FileName = txtIP.Text + "_" + a.Replace(":", "_");
                        saveFileDialog.DefaultExt = "csv";
                        saveFileDialog.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                        // 顯示對話方塊，讓使用者選擇檔案位置及檔名

                        DialogResult result = saveFileDialog.ShowDialog();


                        // 如果使用者按下「儲存」按鈕

                        if (result == DialogResult.OK)
                        {

                            try
                            {
                                button3.Enabled=true;
                                fileName = saveFileDialog.FileName;
                               K=Path.GetDirectoryName(fileName);
                                File.AppendAllLines(fileName, fileContent2);


                                // 將資訊儲存至檔案


                                MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                    }
                    break;
                case 1:
                    {
                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                        String a;
                        Day2 = DateTime.Now;
                        targetTime2 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                        string formattedTime = Day2.ToString("yyyy/MM/dd HH:mm:ss");
                        label6.Text = txtIP.Text + "_" + formattedTime;
                        // 設定預設的檔案副檔名及檔案類型篩選
                        a = formattedTime.Replace("/", ".");
                        saveFileDialog1.FileName = txtIP.Text + "_" + a.Replace(":", "_");

                        saveFileDialog1.DefaultExt = "csv";
                        saveFileDialog1.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                        // 顯示對話方塊，讓使用者選擇檔案位置及檔名
                        DialogResult result = saveFileDialog1.ShowDialog();

                        // 如果使用者按下「儲存」按鈕
                        if (result == DialogResult.OK)
                        {


                            try
                            {
                                button3.Enabled = true;
                                fileName2 = saveFileDialog1.FileName;
                                L = Path.GetDirectoryName(fileName2);
                                File.AppendAllLines(fileName2, fileContent2);


                                // 將資訊儲存至檔案


                                MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    break;
                case 2:
                    {
                        SaveFileDialog saveFileDialog2 = new SaveFileDialog();
                        String a;
                        Day3 = DateTime.Now;
                        targetTime3 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                        string formattedTime = Day3.ToString("yyyy/MM/dd HH:mm:ss");
                        label6.Text = txtIP.Text + "_" + formattedTime;
                        // 設定預設的檔案副檔名及檔案類型篩選
                        a = formattedTime.Replace("/", ".");
                        saveFileDialog2.FileName = txtIP.Text + "_" + a.Replace(":", "_");
                        // 設定預設的檔案副檔名及檔案類型篩選
                        saveFileDialog2.DefaultExt = "csv";
                        saveFileDialog2.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                        // 顯示對話方塊，讓使用者選擇檔案位置及檔名
                        DialogResult result = saveFileDialog2.ShowDialog();

                        // 如果使用者按下「儲存」按鈕
                        if (result == DialogResult.OK)
                        {


                            try
                            {
                                button3.Enabled = true;
                                fileName3 = saveFileDialog2.FileName;
                                G = Path.GetDirectoryName(fileName3);
                                File.AppendAllLines(fileName3, fileContent2);


                                // 將資訊儲存至檔案


                                MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    break;
            }



        }

        private void SaveFile2()
        {
            switch (comboBox1.SelectedIndex)
            {



                case 0:
                    {

                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        String a;
                        Day4 = DateTime.Now;
                        targetTime4 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                        string formattedTime = Day4.ToString("yyyy/MM/dd HH:mm:ss");
                        label7.Text = txtIP.Text + "_" + formattedTime;

                        // 設定預設的檔案副檔名及檔案類型篩選
                        a = formattedTime.Replace("/", ".");
                        saveFileDialog.FileName = txtIP.Text + "_" + a.Replace(":", " ");
                        saveFileDialog.DefaultExt = "csv";
                        saveFileDialog.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                        // 顯示對話方塊，讓使用者選擇檔案位置及檔名

                        DialogResult result = saveFileDialog.ShowDialog();


                        // 如果使用者按下「儲存」按鈕

                        if (result == DialogResult.OK)
                        {

                            try
                            {
                                button9.Enabled = true;
                                fileName = saveFileDialog.FileName;
                                K = Path.GetDirectoryName(fileName);
                                File.AppendAllLines(fileName, fileContent3);


                                // 將資訊儲存至檔案


                                MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                    }
                    break;
                case 1:
                    {
                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                        String a;
                        Day5 = DateTime.Now;
                        targetTime5 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                        string formattedTime = Day5.ToString("yyyy/MM/dd HH:mm:ss");
                        label7.Text = txtIP.Text + "_" + formattedTime;
                        // 設定預設的檔案副檔名及檔案類型篩選
                        a = formattedTime.Replace("/", ".");
                        saveFileDialog1.FileName = txtIP.Text + "_" + a.Replace(":", " ");
                        saveFileDialog1.DefaultExt = "csv";
                        saveFileDialog1.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                        // 顯示對話方塊，讓使用者選擇檔案位置及檔名
                        DialogResult result = saveFileDialog1.ShowDialog();

                        // 如果使用者按下「儲存」按鈕
                        if (result == DialogResult.OK)
                        {


                            try
                            {

                                button9.Enabled = true;
                                fileName2 = saveFileDialog1.FileName;
                                L = Path.GetDirectoryName(fileName2);
                                File.AppendAllLines(fileName2, fileContent3);


                                // 將資訊儲存至檔案


                                MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    break;
                case 2:
                    {
                        SaveFileDialog saveFileDialog2 = new SaveFileDialog();
                        String a;
                        Day6 = DateTime.Now;
                        targetTime6 = DateTime.Today.AddHours(23).AddMinutes(59).AddSeconds(59); // 添加一天
                        string formattedTime = Day6.ToString("yyyy/MM/dd HH:mm:ss");
                        label7.Text = txtIP.Text + "_" + formattedTime;
                        // 設定預設的檔案副檔名及檔案類型篩選
                        a = formattedTime.Replace("/", ".");
                        saveFileDialog2.FileName = txtIP.Text + "_" + a.Replace(":", " ");
                        // 設定預設的檔案副檔名及檔案類型篩選
                        saveFileDialog2.DefaultExt = "csv";
                        saveFileDialog2.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                        // 顯示對話方塊，讓使用者選擇檔案位置及檔名
                        DialogResult result = saveFileDialog2.ShowDialog();

                        // 如果使用者按下「儲存」按鈕
                        if (result == DialogResult.OK)
                        {


                            try
                            {
                                button9.Enabled = true;
                                fileName3 = saveFileDialog2.FileName;
                                G = Path.GetDirectoryName(fileName3);
                                File.AppendAllLines(fileName3, fileContent3);


                                // 將資訊儲存至檔案


                                MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    break;
            }



        }
        //  string[] a =

        //{  "V_a,4352,float",
        //  "I_a,4354,float",
        //  "kW_a,4356,float",
        //  "kVAR_a,4358,float",
        //  "kVA_a,4360,float",
        //  "PF_a,4362,float",
        //  "kWh_a,4364,float",
        //  "kVARh_a,4366,float",
        //  "kVAh_a,4368,float",
        //  "V_b,4370,float",
        //  "I_b,4372,float",
        //  "kW_b,4374,float",
        //  "kVAR_b,4376,float",
        //  "kVA_b,4378,float",
        //  "PF_b,4380,float",
        //  "kWh_b,4382,float",
        //  "kVARh_b,4384,float",
        //  "kVAh_b,4386,float",
        //  "V_c,4388,float",
        //  "I_c,4390,float",
        //  "kW_c,4392,float",
        //  "kVAR_c,4394,float",
        //  "kVA_c,4396,float",
        //  "PF_c,4398,float",
        //  "kWh_c,4400,float",
        //  "kVARh_c,4402,float",
        //  "kVAh_c,4404,float",
        //  "V_avg,4406,float",
        //  "I_avg,4408,float",
        //  "W_tot,4410,float",
        //  "VAR_tot,4412,float",
        //  "VA_tot,4414,float" ,
        //  "PF_tot,4416,float",
        //  "kWh_tot,4418,float",
        //  "kVARh_tot,4420,float",
        //  "kVAh_tot,4422,float",
        //  "Freq_a,4424,float",
        //  "Freq_b,4426,float",
        //  "Freq_c,4428,float",
        //  "Freq_tot,4430,float"

        //           };
        //  for (int i = 0; i < a.Length; i++)
        //  {
        //      string[] ar = a[i].Split(',');
        //      //int[] value = plc.ReadHoldingRegisters(Convert.ToInt32(ar[1]), 2); ;
        //      dataGridView1.Rows.Add(ar[0], ar[1], ar[2],String.Empty);


        //  }



        //for (int i = 4352; i <= 4430; i += 2)
        //{
        //    dataGridView1.Rows.Add(i.ToString(), "float", string.Empty, string.Empty);
        //}
        //foreach (DataGridViewRow Row2 in dataGridView1.Rows)
        //{



        //}
    }
}

