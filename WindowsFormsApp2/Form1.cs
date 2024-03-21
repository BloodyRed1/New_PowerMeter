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
using System.Diagnostics;
using System.Data;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        DateTime one;
        DateTime two;
        int Line=0;
        string K,V;
        int lastid = 1;
        string fileName;
        DateTime startTime;
        List<List<string>> dataList = new List<List<string>>();
        string[] ar;
        int AIndex = 0;
        int BIndex = 0;
        int totalvaule = 0;
        bool titlewrite=false;
        List<float> modbusHZ = new List<float>();
        List<string> ReadTime = new List<string>();
        List<string> pingspeed = new List<string>();
        private CancellationTokenSource cancellationTokenSource2;
        private CancellationTokenSource cancellationTokenSource3;
        ModbusClient plc = new ModbusClient();
        bool Unconnect = false;
        int disconnect_count = 0;
        bool x = false;
        List<string> fileContent2 = new List<string>();
        List<string> fileContent = new List<string>();
        List<string> fileContent3 = new List<string>();
        List<string> fileContent4 = new List<string>();
        List<string> fileContent5 = new List<string>();
        string IP1;
        short port;
        double timeout = 0;
        int timeout2 = 0;
        int timeout3 = 0;
        int x2 = 0;
        int start = 0;
        private Thread plcThread;
        private static readonly object fileLock = new object();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

            }
        }

        private void btmConnect_Click(object sender, EventArgs e)
        {

            int aa = 0;
            try
            {
                timeout2 = Convert.ToInt32(txtTimeout.Text.Trim());
                timeout = Convert.ToInt32(txtTimeout.Text.Trim());
                timeout3 = 200 - timeout2;
                //timeout = 1 / timeout*1000000;

                timeout = 1 / timeout * 1000;//將分數x1000單位加上m變成小數
                bool isPing = PingHost(txtIP.Text.Trim(), timeout);

                if (isPing == false)
                {
                    MessageBox.Show("無法連接", "錯誤", MessageBoxButtons.OK);
                    return;
                }


                else
                {
                    if (Unconnect == false)
                    {
                        Unconnect = true;
                        plc?.Disconnect();
                        IP1 = txtIP.Text.Trim();
                        port = Convert.ToInt16(txtPort.Text.Trim());
                        plc.IPAddress = IP1;
                        plc.Port = port;
                        plc.Connect();
                        setStatusLabel();



                        //if (dataGridView1.RowCount <= 0)
                        ////{
                        //    dataGridView1.Rows.Clear();


                        string line = string.Empty;
                        int rr = 0;
                        using (StreamReader sr = new StreamReader(textBox1.Text))
                        {
                            while ((line = sr.ReadLine()) != null)
                            {
                                ar = line.Split(',');
                                if (rr > 0) dataList.Add(new List<string> { aa.ToString(), ar[1], ar[2], ar[3], String.Empty });
                                //if (rr > 0) dataGridView1.Rows.Add(aa.ToString(), ar[1], ar[2], ar[3], String.Empty);
                                rr++;
                                aa += 1;
                            }
                            Line = rr - 1;
                        }
                        //}
                        btmConnect.Text = "中斷";
                        if (plc.Connected == true)
                            MessageBox.Show("PLC連線成功。", "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        label1.Text = txtIP.Text.Trim();
                    }


                    //}

                    else
                    {
                        if (plc.Connected == true)
                        {

                            MessageBox.Show("連線失敗次數" + disconnect_count, "訊息！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (x2 > 0)
                            {
                                //cancellationTokenSource2.Cancel();
                                cancellationTokenSource3.Cancel();
                            }

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
                string msg = $"Exception occurred: {ex.Message}";
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                WriteToLogFile(msg);


            }
            finally
            {

            }
        }

        public static bool PingHost(string _ip, double Timeout) //判斷IP是否還可以連線
        {

            bool pingable = false;
            Ping pinger = null;

            try
            {

                pinger = new Ping();
                PingReply reply = pinger.Send(_ip, timeout: (int)Timeout);
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

        public static bool PingHost2(string _ip, int Timeout) //判斷IP是否還可以連線
        {

            bool pingable = false;
            Ping pinger = null;

            try
            {
                pinger = new Ping();
                pingable = pinger.Send(_ip, timeout: Timeout).Status == IPStatus.Success;
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

        private void setStatusLabel()
        {
            bool connect = (plc == null) ? false : plc.Connected;
            Lbstatus.Text = string.Format("連線狀態： {0}", connect);
        }
        private void ReadPLC3()
        {
            one = DateTime.Now.AddMinutes(1);
            // 創建執行緒
            cancellationTokenSource3 = new CancellationTokenSource();
            plcThread = new Thread(new ThreadStart(() =>
            {
                while (true)
                {
                    // 檢查取消請求
                    if (cancellationTokenSource3.Token.IsCancellationRequested)
                    {
                        plcThread.Join();
                        // 如果取消，則退出循環
                        break;
                    }
                    try
                    {
                        string plcIpAddress = txtIP.Text.Trim(); 
 
                       
                         // 建立一個 List 來存儲所有的 Modbus 通訊結果
                            List<int[]> modbusResults = new List<int[]>();
                            int rowIndex = 0;
                            if (AIndex < totalvaule)
                            {
                                foreach (List<string> item in dataList)
                                {
                                
                                    if (item.Count > 2)
                                    {
                                        int address = Convert.ToInt32(item[2]);
                                        int[] x2 = plc.ReadHoldingRegisters(address, 2);
                                        modbusResults.Add(x2);
                                    }
                                    int[] x = modbusResults[rowIndex];
                                    float aa = ModbusClient.ConvertRegistersToFloat(x);
                                        try
                                        {
                                            modbusHZ.Add(aa);
                                            AIndex++;
                                       // if (AIndex % timeout2 == 0)
                                      //  {
                                            // Thread.Sleep(timeout3);
                                      //  }
                                        rowIndex++;
                                        }
                                       catch (Exception ex)
                                       {
                                            //string xx = ex.Message;
                                            //WriteToLogFile(xx);
                                       }
                                }
                                DateTime currentTime = DateTime.Now;
                                string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss:fff");
                                two =DateTime.Now;
                                ReadTime.Add(formattedTime);
                            List<string> columnNames = new List<string> { "ID", "DateTime", "V_AB", "I_a", "V_BC", "I_b", "V_CA", "I_c", "LineVolt_avg", "I_avg", "kWh", "kVARh", "kVAh", "KW_tot", "KVAR_tot", "KVA_tot", "PF_tot", "DPF_tot", "Freq_tot" };
                            if (two >= one)
                                {
                                    currentTime = DateTime.Now;
                                    formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");
                                    string f = "";
                                    f = formattedTime.Replace("/", ".");
                                    fileName = K + "\\" + V + "_" + f.Replace(":", " ") + ".csv";
                                    one = DateTime.Now.AddMinutes(1);
                                    titlewrite = false;
                                      int S=0;
                                WriteListToFile(modbusHZ, fileName, Line, columnNames, ReadTime, titlewrite, S ,two,one);
                            }
                              
                                bool x1=false;
                            if (x1 == false)
                            {
                                int S= 0;
                                
                                x1= true;
                            }




                            titlewrite = true;
                            }
                            if (AIndex >= totalvaule)
                            {
                                this.Invoke((MethodInvoker)delegate
                                {
                                    label5.Text = AIndex.ToString();
                                    label4.Text = "存取完成";
                                });
                                plcThread.Join(); 
                            }
                        }
                    
                    catch (Exception ex)
                    {
                      
                            bool isPing = PingHost(txtIP.Text.Trim(), timeout);
                            Ping pingSender = new Ping();
                            if (isPing == false)
                            {
                                plc?.Disconnect();
                                plc.Connect();
                                //string xx = ex.Message;
                                //WriteToLogFile(xx);

                            }
               
                    }
                   
                    if (cancellationTokenSource3.Token.IsCancellationRequested)
                    {
                        plcThread.Join();
                        // 如果取消，則退出循環
                        break;
                    }
                }
            }));
            // 啟動執行緒
            plcThread.Start();
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
              
            }
        }
        //if (AIndex % 5000 == 0)  // 每5000次更新一次UI
        //{
        //    this.Invoke((MethodInvoker)delegate
        //    {
        //        label5.Text = "目前數量:" + AIndex.ToString();
        //    });
        //}
        //    private void ReadPLC()
        //    {
        //        int aaa;
        //        float aa=0;

        //        cancellationTokenSource2 = new CancellationTokenSource();

        //            ThreadPool.QueueUserWorkItem(async state =>
        //            {

        //                await Task.Run(() =>
        //                {

        //                    while (true)
        //                    {
        //                        // 檢查取消請求
        //                        if (cancellationTokenSource2.Token.IsCancellationRequested)
        //                        {
        //                            // 如果取消，則退出循環
        //                            break;
        //                        }

        //                        try
        //                        {

        //                            bool isPing = PingHost(txtIP.Text.Trim(), timeout);
        //                            if (isPing == false)
        //                            {
        //                                disconnect_count += 1;
        //                                //Lbstatus.Text = string.Format("連線狀態： {0}", "連線中斷");
        //                                plc?.Disconnect();
        //                                IP1 = txtIP.Text.Trim();
        //                                port = Convert.ToInt16(txtPort.Text.Trim());
        //                                plc.Connect();
        //                                //Lbstatus.Text = string.Format("連線狀態： {0}", "再次連線");
        //                            }
        //                            else
        //                            {
        //                                DateTime currentTime = DateTime.Now;
        //                                string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");
        //                                // 建立一個 List 來存儲所有的 Modbus 通訊結果
        //                                List<int[]> modbusResults = new List<int[]>();
        //                                // 第一次迴圈：進行 Modbus 通訊的讀取操作，並將結果存入 modbusResults
        //                                foreach (DataGridViewRow Row in dataGridView1.Rows)
        //                                {
        //                                    int address = Convert.ToInt32(Row.Cells[2].Value);
        //                                    int[] x = plc.ReadHoldingRegisters(address, 2);
        //                                    modbusResults.Add(x);
        //                                }
        //                                // 第二次迴圈：將讀取到的 Modbus 資料分配給每一列的儲存格
        //                                int rowIndex = 0;
        //                                foreach (DataGridViewRow Row in dataGridView1.Rows)
        //                                {
        //                                    int[] x = modbusResults[rowIndex];
        //                                    // 假設要將結果存入儲存格的第四個欄位（即 Cells[3]）
        //                                    aa = ModbusClient.ConvertRegistersToFloat(x);
        //                                    Row.Cells[4].Value = aa.ToString();
        //                                    if(AIndex<=totalvaule)
        //                                    {
        //                                        modbusHZ.Add(aa);
        //                                        AIndex++;
        //                                        try
        //                                        {
        //                                            this.Invoke((MethodInvoker)delegate
        //                                            {
        //                                                label5.Text = "目前數量:"+AIndex.ToString();

        //                                            });
        //                                        }
        //                                        catch (Exception ex)
        //                                        {
        //                                            string xx = ex.Message;
        //                                            WriteToLogFile(xx);
        //                                        }
        //                                        //if (modbusHZ.Count <= totalvaule)
        //                                        //{
        //                                        //}
        //                                        //else
        //                                        //{
        //                                        //   
        //                                        //}
        //                                    }
        //                                    else
        //                                    {
        //                                        try
        //                                        {
        //                                            this.Invoke((MethodInvoker)delegate
        //                                            {
        //                                                label1.Text = "数据收集完成";
        //                                            });
        //                                        }
        //                                        catch (Exception ex)
        //                                        {
        //                                            string xx = ex.Message;
        //                                            WriteToLogFile(xx);
        //                                        }

        //                                    }

        //                                    //switch (rowIndex)
        //                                    //{
        //                                    //    case 0:

        //                                        //        break;
        //                                        //    case 1:
        //                                        //        break;
        //                                        //    case 2:
        //                                        //        break;
        //                                        //}

        //                                        rowIndex++;
        //                                }

        //                                // 更新顯示
        //                                dataGridView1.Invalidate();







        //                                Thread.Sleep((int)timeout);
        //                                //DelayMicroseconds(timeout);

        //                            }
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            try
        //                            {
        //                                string xx = ex.Message;
        //                                WriteToLogFile(xx);
        //                            }
        //                            catch { ex.ToString(); }
        //                        }

        //                        if (cancellationTokenSource2.Token.IsCancellationRequested)
        //                        {
        //                            // 如果取消，則退出循環
        //                            break;
        //                        }
        //                    }
        //                }, cancellationTokenSource2.Token);
        //            });


        //}
        //    private void ReadPLC2()
        //    {
        //        //DateTime endTime = DateTime.Now; // 当前时间
        //        //TimeSpan elapsedTime = endTime - startTime;

        //        //// 获取分钟和秒
        //        //int minutes = (int)elapsedTime.TotalMinutes;
        //        //int seconds = (int)elapsedTime.TotalSeconds;

        //        //Console.WriteLine($"已经过去了 {minutes} 分钟和 {seconds} 秒。");
        //        float aa = 0;

        //        cancellationTokenSource3 = new CancellationTokenSource();

        //        ThreadPool.QueueUserWorkItem(async state =>
        //        {

        //            await Task.Run(async() =>
        //            {

        //                while (true)
        //                {
        //                    // 檢查取消請求
        //                    if (cancellationTokenSource3.Token.IsCancellationRequested)
        //                    {
        //                        // 如果取消，則退出循環
        //                        break;
        //                    }
        //                    try
        //                    {
        //                        bool isPing = PingHost(txtIP.Text.Trim(), timeout);
        //                        if (isPing == false)
        //                        {

        //                            plc?.Disconnect();
        //                            plc.Connect();
        //                        }

        //                        //}
        //                        //else
        //                        //{
        //                        // 建立一個 List 來存儲所有的 Modbus 通訊結果
        //                        List<int[]> modbusResults = new List<int[]>();
        //                            // 第一次迴圈：進行 Modbus 通訊的讀取操作，並將結果存入 modbusResults
        //                            //foreach (DataGridViewRow Row in dataGridView1.Rows)
        //                            //{
        //                            //    int address = Convert.ToInt32(Row.Cells[2].Value);
        //                            //    int[] x = plc.ReadHoldingRegisters(address, 2);
        //                            //    modbusResults.Add(x);
        //                            //}
        //                            // 第二次迴圈：將讀取到的 Modbus 資料分配給每一列的儲存格
        //                            int rowIndex = 0;
        //                        //foreach (DataGridViewRow Row in dataGridView1.Rows)
        //                        //{
        //                        if (AIndex < totalvaule)
        //                        { 

        //                            foreach (List<string> item in dataList)
        //                            {

        //                                if (item.Count > 2)
        //                                {
        //                                    //ar2Values.Add(item[2]);
        //                                    int address = Convert.ToInt32(item[2]);
        //                                    int[] x2 = plc.ReadHoldingRegisters(address, 2);
        //                                    modbusResults.Add(x2);
        //                                }
        //                                int[] x = modbusResults[rowIndex];
        //                                // 假設要將結果存入儲存格的第四個欄位（即 Cells[3]）
        //                                aa = ModbusClient.ConvertRegistersToFloat(x);
        //                                    try
        //                                    {
        //                                    //if (AIndex < totalvaule)
        //                                    //{
        //                                    modbusHZ.Add(aa);
        //                                    AIndex++;


        //                                    if (AIndex == totalvaule)
        //                                    {
        //                                            this.Invoke((MethodInvoker)delegate
        //                                               {
        //                                                   label5.Text = AIndex.ToString();
        //                                                   label4.Text = "抓取完成";
        //                                               });
        //                                        }
        //                                        if (AIndex % 5000 == 0)  // 每5000次更新一次UI
        //                                        {
        //                                            this.Invoke((MethodInvoker)delegate
        //                                            {
        //                                                label5.Text = "目前數量:" + AIndex.ToString();
        //                                            });
        //                                        }
        //                                    }
        //                                    catch (Exception ex )
        //                                    {
        //                                        string xx = ex.Message;
        //                                        WriteToLogFile(xx);
        //                                    }          
        //                                    rowIndex++;

        //                            }
        //                            //if (AIndex % timeout2 == 0)
        //                            //{
        //                            //    await Task.Delay((int)timeout); 
        //                            //}
        //                            DateTime currentTime = DateTime.Now;
        //                            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss:fff");
        //                            ReadTime.Add(formattedTime);
        //                        }
        //                        //DelayMicroseconds(timeout);
        //                        //}
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        try
        //                        {
        //                            string xx = ex.Message;
        //                            WriteToLogFile(xx);
        //                        }
        //                        catch { ex.ToString(); }
        //                    }

        //                    if (cancellationTokenSource3.Token.IsCancellationRequested)
        //                    {
        //                        // 如果取消，則退出循環
        //                        break;
        //                    }
        //                }
        //            } , cancellationTokenSource3.Token);
        //        });


        //    }
        public void WriteToLogFile(string message)
        {
            string logFileName = "Logs\\ErrorFile2.txt"; // 指定日志文件的路径
            string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logFileName);
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                // 写入信息到文件
                writer.WriteLine(DateTime.Now + ": " + message);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            V = label1.Text;
            startTime = DateTime.Now; // 你的起始时间
            if (x == false)
            {
                try
                {
                    lastid = 1;
                    totalvaule = Convert.ToInt32(textBox2.Text);
                    totalvaule = totalvaule * Line;
                }
                catch (Exception ex)
                {
                  
                }
                button2.Text = "停止抓取";
                x2 += 1;
                //ReadPLC();
                //ReadPLC2();
                ReadPLC3();
                x = true;
                label4.Text = "抓取中";
            }
            else
            {
                button2.Text = "抓取";
                x = false;
               
                cancellationTokenSource3.Cancel();
                
            }


        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // 觸發 MainForm 的 Show() 方法
                (Owner as Form3)?.Show();
            }
          
        }

        private void button3_Click(object sender, EventArgs e)
        {

            //int x = new int();
            //try
            //{
            //    totalvaule = Convert.ToInt32(textBox2.Text);


            //    string[] ar = new string[1000000];
            //    for (int i = 0; i <= totalvaule; i++)
            //    {
            //        ar[i] = dataGridView2.Rows[i].Cells[4].Value.ToString();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    WriteToLogFile(ex.ToString());
            //}

            // 遍历列表并逐一添加到数据表中
            //for (int i = 0; i < modbusHZ.Count; i++)
            //{
            //    dataGridView2.Rows.Add(String.Empty, String.Empty, String.Empty, String.Empty, String.Empty);
            //    foreach (DataGridViewRow Row in dataGridView2.Rows)
            //    {
            //         dataGridView2.Rows[Row.Index].Cells[0].Value = i + 1;  // 第一个单元格（ID）
            //         dataGridView2.Rows[Row.Index].Cells[4].Value = modbusHZ[i];  // 第二个单元格（Value）
            //    } 

            //}
        }
        //dataGridView2.DataSource = yourDataSource; // 替换为你的数据源
        //private void dataGridView2_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        //{
        //    // 根据需要提供数据，可以根据行号和列号动态生成数据
        //    // 注意：这里的逻辑需要根据实际需求和数据结构进行修改
        //    e.Value = YourDataFetchingLogic(e.RowIndex, e.ColumnIndex);
        //}
        //private object YourDataFetchingLogic(int rowIndex, int colIndex)
        //{
        //    // 根据行号和列号获取相应的数据
        //    // 示例逻辑，你需要根据实际需求修改
        //    return yourDataSource[rowIndex][colIndex];
        //}
        private void button5_Click(object sender, EventArgs e)
        {
          
            SaveInformation();


        }
        private void SaveInformation()
        {
            string a;
            DateTime Day = DateTime.Now;
            string formattedTime = Day.ToString("yyyy/MM/dd HH:mm:ss");
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            // 設定預設的檔案副檔名及檔案類型篩選
            a = formattedTime.Replace("/", ".");
            saveFileDialog.FileName = txtIP.Text + "_" + a.Replace(":", " ");
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.Filter = "csv檔案 (*.csv)|*.csv|xlsx檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

            // 顯示對話方塊，讓使用者選擇檔案位置及檔名

            DialogResult result = saveFileDialog.ShowDialog();
            
            if (result == DialogResult.OK)
            {
                button2.Enabled = true;

                try
                {
                  
                    List<string> columnNames = new List<string> { "ID", "DateTime", "V_AB", "I_a", "V_BC", "I_b", "V_CA", "I_c", "LineVolt_avg", "I_avg", "kWh", "kVARh", "kVAh", "KW_tot", "KVAR_tot", "KVA_tot", "PF_tot", "DPF_tot", "Freq_tot" };
                    fileName = saveFileDialog.FileName;
                    K = Path.GetDirectoryName(fileName);
                    using (StreamWriter writer = File.CreateText(fileName)) // 使用 CreateText 方法創建新檔案
                    { 
                   
                    }
                  
                    MessageBox.Show("檔案已儲存成功！", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("無法儲存檔案：" + ex.Message, "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
            }
        }

        //static async Task DelayMicroseconds(int microseconds)
        //{
        //    // 将微秒转换为 TimeSpan
        //    TimeSpan delayTime = TimeSpan.FromMilliseconds(microseconds / 1000.0);

        //    // 使用 Task.Delay 进行等待
        //    await Task.Delay(delayTime);
        //}

        static void WriteListToFile<T>(List<T> dataList, string filePath, int valuesPerLine, List<string> columnNames, List<string> readTimeList ,bool title, int lastId  ,DateTime ss ,DateTime vv)
        {
            try
            {
                lock (fileLock)
                {
                    // 使用 StreamWriter 将数据写入文件
                    using (StreamWriter writer = new StreamWriter(filePath, true)) // 使用 true 表示追加到现有文件的末尾
                    {
                        if (!title) // 如果 title 为 false，则写入标题
                        {
                            string header = string.Join(", ", columnNames);
                            writer.WriteLine(header);
                        }

                        // 将数据列表转换为文本格式
                        List<string> dataLines = new List<string>();
                        for (int i = 0; i < dataList.Count; i++)
                        {
                            string line = (lastId++) + ", " + readTimeList[i] + ", "; // 添加编号和时间戳
                            line += string.Join(", ", dataList[i].ToString());
                            dataLines.Add(line);
                        }

                        // 将整个文本写入文件
                        if (ss > vv)
                        {
                            foreach (string line in dataLines)
                            {
                                writer.WriteLine(line);
                            }
                        }
                      
                    }
                }

                // 清除 dataList、columnNames 和 readTimeList，以便下一次调用时使用新的数据
                dataList.Clear();
                columnNames.Clear();
                readTimeList.Clear();
                
            }
            catch (Exception ex)
            {
                // 异常处理逻辑
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            plc?.Disconnect();
            cancellationTokenSource3?.Dispose();
            plcThread?.Abort();
        }
    }

}
