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
using System.Net.Sockets;
namespace WindowsFormsApp2
{

    public partial class Form5 : Form
    {

        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            button3.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

            }
        }
        //匯出 csv
        public static void exportToCSV(StringBuilder sb, int count, string title)
        {
            //存檔
            SaveFileDialog csvFile = new SaveFileDialog();
            csvFile.Filter = "csv檔|*.csv";
            csvFile.Title = title + "_另存新檔";
            //預設檔名
            csvFile.FileName = string.Format("{0}{2}匯出-{1}", DateTime.Today.ToString("yyyyMMdd"), count, title);
            csvFile.DefaultExt = "*.csv";
            if (csvFile.ShowDialog() == DialogResult.OK)
            {
                System.IO.File.WriteAllText(csvFile.FileName.ToString(), sb.ToString());
            }


            //釋放
            sb = null;


        }
        private async void btmConnect_Click(object sender, EventArgs e)
        {
            try
            {
                var tcpClient = new TcpClient();
                await tcpClient.ConnectAsync("172.16.178.5", 502);
                //发送请求封包
                Int16 address = 264;                           /*需要指定为16位, 我这里默认是32位, 否则reverse后是16#0000 07d0, 返回0x83*/
                Int16 quantity = 1;
                var startingAddress = BitConverter.GetBytes(address);
                var quantityOfRegisters = BitConverter.GetBytes(quantity);

                if (BitConverter.IsLittleEndian)
                {
                    startingAddress = startingAddress.Reverse().ToArray();
                    quantityOfRegisters = quantityOfRegisters.Reverse().ToArray();
                }
                var request = new byte[]
                    {
                      0x00,0x01,                                       //Transaction Identifiter, 连续数字或随机数字都可以
                      0x00,0x00,                                       //Protocol Identifier, 固定0
                      0x00,0x06,                                       //Length, 接下来有6byte的资料
                      0x01,                                            //Unit Identifier
                      0x03,                                            //Function Code
                        startingAddress[0],startingAddress[1],           //Starting address (2002转成2 bytes的结果)
                        quantityOfRegisters[0],quantityOfRegisters[1]    //Quantity of Registers (1转成2bytes的结果)
                                                                         };
               textBox3.Text+=("发送完毕.");
               textBox3.Text+=(BitConverter.ToString(request));

                //获取网络流
                var stream = tcpClient.GetStream(); /*原文没实例化*/
                //发送数据
                await stream.WriteAsync(request, 0, request.Length);

                var response = new byte[1024];
                var readCount =  await stream.ReadAsync(response, 0, response.Length);
                // 比較完整的寫法是把 Transaction Identifier、Function Code 等都抓出來檢查
                // 尤其是 Function Code，因為當有 error 時就不會是 0x03，而是 0x83 了
                // 不過這邊就先省略，以抓到我們的目標資料為主，那麼就是最後的 quantity * 2 個 bytes
                //var result = response.Take(readCount).TakeLast(quantity * 2).ToArray();

                textBox3.Text += ("接收完毕");
                textBox3.Text += (BitConverter.ToString(response));
                textBox3.Text += ("读取数量:" + readCount);

            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误:" + ex.Message);

            }
        }
    }
}