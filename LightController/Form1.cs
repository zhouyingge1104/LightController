using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net.Sockets;
using System.Threading;
using System.Timers;

namespace LightController
{
    public partial class FormMain : Form
    {
        const int WHITE = 0, RED = 1, GREEN = 2, BLUE = 3, YELLOW = 4, ORANGE = 5;
        const int SETTING_CONTENT_LENGTH = 359;

        int page; //当前页码
        string crcCheck;
 
        uint Test;
        private delegate void InvokeCallback(string msg); //定义回调函数（代理）格式

        System.Timers.Timer timer;

        public FormMain()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;//关闭跨线程调用检测 
            
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            this.Width = Convert.ToInt32(Screen.GetWorkingArea(this).Width * 0.9);
            this.Height = Convert.ToInt32(Screen.GetWorkingArea(this).Height * 0.7);
            this.CenterToScreen();

            cbxPage.SelectedIndex = 0;

            //for test
            byte[] a = { 1, 2, 3 };
            byte[] b = new byte[6];
            a.CopyTo(b, 0);
            string str = "";
            for (int i = 0; i < b.Length; i ++ )
            {
                str += b[i] + " ";
            } 
            //MessageBox.Show(str);
           
            dgv.RowCount = 12;
            dgv.Rows[0].HeaderCell.Value = "功能1"; 
            dgv.Rows[1].HeaderCell.Value = "功能2";
            dgv.Rows[2].HeaderCell.Value = "功能3";
            dgv.Rows[3].HeaderCell.Value = "功能4";
            dgv.Rows[4].HeaderCell.Value = "功能5";
            dgv.Rows[5].HeaderCell.Value = "功能6";
            dgv.Rows[6].HeaderCell.Value = "功能7";
            dgv.Rows[7].HeaderCell.Value = "功能8";
            dgv.Rows[8].HeaderCell.Value = "功能9";
            dgv.Rows[9].HeaderCell.Value = "功能10";
            dgv.Rows[10].HeaderCell.Value = "功能11";
            dgv.Rows[11].HeaderCell.Value = "功能12";

            dgv.Rows[0].Cells[0].Selected = false;

            getSavedSetting();
            getSavedSettingNet();

            timer = new System.Timers.Timer();
            timer.Elapsed += new ElapsedEventHandler(TimerTask);
            timer.Interval = 3000;
            timer.AutoReset = false;

        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {

           int left = dgv.Columns[0].Width * (dgv.CurrentCell.ColumnIndex + 2) + this.Left;
           int top = dgv.Rows[0].Height * (dgv.CurrentCell.RowIndex + 3) + this.Top;
             
           colors.Show();
           colors.SetBounds(left, top, 100, 400);
          
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dgv.CurrentCell.Style.BackColor = Color.White;
            dgv.CurrentCell.Style.ForeColor = Color.White;
            dgv.CurrentCell.Value = WHITE;
            dgv.CurrentCell.Selected = false;
            colors.Hide();

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            dgv.CurrentCell.Style.BackColor = Color.Red;
            dgv.CurrentCell.Style.ForeColor = Color.Red;
            dgv.CurrentCell.Value = RED;
            dgv.CurrentCell.Selected = false;
            colors.Hide();
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            dgv.CurrentCell.Style.BackColor = Color.Green;
            dgv.CurrentCell.Style.ForeColor = Color.Green;
            dgv.CurrentCell.Value = GREEN;
            dgv.CurrentCell.Selected = false;
            colors.Hide();
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            dgv.CurrentCell.Style.BackColor = Color.Blue;
            dgv.CurrentCell.Style.ForeColor = Color.Blue;
            dgv.CurrentCell.Value = BLUE;
            dgv.CurrentCell.Selected = false;
            colors.Hide();
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            dgv.CurrentCell.Style.BackColor = Color.Yellow;
            dgv.CurrentCell.Style.ForeColor = Color.Yellow;
            dgv.CurrentCell.Value = YELLOW;
            dgv.CurrentCell.Selected = false;
            colors.Hide();
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            dgv.CurrentCell.Style.BackColor = Color.Orange;
            dgv.CurrentCell.Style.ForeColor = Color.Orange;
            dgv.CurrentCell.Value = ORANGE;
            dgv.CurrentCell.Selected = false;
            colors.Hide();
        }


        private void colors_Closing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            dgv.CurrentCell.Selected = false;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            save();
        }

        private void btnSaveAndSend_Click(object sender, EventArgs e)
        {
            tbxMsg.Text = "";
            Thread.Sleep(1000);
            save();
            send();
        }

        /// <summary>
        /// 保存
        /// </summary>
        private void save()
        {
            String currPath = Environment.CurrentDirectory;
            String fileName = "L" + page + ".txt";
            String fileNameNet = "net.txt";
            FileInfo setting = new FileInfo(currPath + "/" + fileName);
            if (!setting.Exists)
            {
                setting.Create().Close();
            }

            FileInfo settingNet = new FileInfo(currPath + "/" + fileNameNet);
            if (!settingNet.Exists)
            {
                settingNet.Create().Close();
            }
            StreamWriter writer1 = new StreamWriter(setting.FullName);
            writer1.Write(getCurrSetting());

            StreamWriter writer2 = new StreamWriter(settingNet.FullName);
            writer2.Write(getCurrSettingNet());

            writer1.Dispose();
            writer2.Close();

            writer1.Dispose();
            writer2.Close();

            tbxMsg.Text = "第 " + page + " 页 保存成功";
        }

        /// <summary>
        /// 发送
        /// </summary>
        private void send()
        {
            tbxMsg.Text = null;
            btnSaveAndSend.Enabled = false;

            byte[] fullOrder = new byte[186]; // 4 + 180 + 2
            byte[] colorOrder = new byte[180]; // 4 + 180

            fullOrder[0] = 0xFF;
            fullOrder[1] = 0x5A;
            fullOrder[2] = (byte)page;
            fullOrder[3] = 0xBA;
            //1. 组织byte[]数组
            int i = 0;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                int j = 0;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    colorOrder[15 * i + j] = byte.Parse(cell.Value.ToString());
                    j ++;
                }
 
                i ++;
            }

            colorOrder.CopyTo(fullOrder, 4);

            byte[] tempOrder = new byte[184];
            tempOrder[0] = 0xFF;
            tempOrder[1] = 0x5A;
            tempOrder[2] = (byte)page;
            tempOrder[3] = 0xBA;
            colorOrder.CopyTo(tempOrder, 4);


            byte[] crcResult = CRC16MODBUS(tempOrder, (byte)tempOrder.Length);


            fullOrder[184] = crcResult[0];
            fullOrder[185] = crcResult[1];

            crcCheck = crcResult[0].ToString("x2").ToUpper() + crcResult[1].ToString("x2").ToUpper();
          
            TcpClient tcp = new TcpClient();
            tcp.SendTimeout = 5000;
            tcp.ReceiveTimeout = 8000;
            try
            {
                 tcp.Connect(tbxIp.Text, 8899);
            }catch(Exception e){
               
                MessageBox.Show("无法连接：" + e.Message);
                btnSaveAndSend.Enabled = true;
                return;
            }

            if (tcp.Connected)
            {
                
                NetworkStream stream = tcp.GetStream();
                stream = tcp.GetStream();//创建于服务器连接的数据流
              
               // stream.BeginWrite(fullOrder, 0, fullOrder.Length, new AsyncCallback(SendCallback), stream);//异步发送数据
                stream.ReadTimeout = 8000;
                stream.Write(fullOrder, 0, fullOrder.Length);
                
                string result = "";
                try
                {
                    
                    byte[] buff = new byte[2];
                    //MessageBox.Show("Begin reading...");
                    stream.Read(buff, 0, 2);
                    //MessageBox.Show("End reading...");
                    string temp = buff[0].ToString("x2").ToUpper() + buff[1].ToString("x2").ToUpper();
                    if (temp.Equals(crcCheck))
                    {
                        result = "发送成功";
                    }
                    else
                    {
                        result = "返回信息不正确，预期：" + crcCheck.ToUpper() + "   实际：" + temp.ToUpper();
                    }

                }
                catch (Exception e)
                {
                    stream.Dispose();
                    stream.Close();

                    tcp.Close();

                    result = "读取失败";
                    MessageBox.Show("无法读取：" + e.Message);
                }

                tbxMsg.Text = result;
                btnSaveAndSend.Enabled = true;

                stream.Dispose();
                stream.Close();

               if(tcp != null && tcp.Connected){
                   tcp.Close();
               }

                timer.Start();

            }
          }

        /// <summary>
        /// Timer任务
        /// </summary>
        private void TimerTask(object  source, ElapsedEventArgs e)
        {
            tbxMsg.Text = null;
            timer.Stop();
        }

        /// <summary>
        /// invoke回调函数
        /// </summary>
        /// <param name="text"></param>
        public void UpdateText(string text)
        {
            if (tbxMsg.InvokeRequired)//当前线程不是创建线程 
                tbxMsg.Invoke(new InvokeCallback(UpdateText), new object[] { text });//回调 
            else//当前线程是创建线程（界面线程） 
                tbxMsg.Text = text;//直接更新 
                btnSaveAndSend.Enabled = true;
        }
        
        /// <summary>
        /// 获取当前的灯光设置
        /// </summary>
        private string getCurrSetting()
        {
            string setting = "";
           
            foreach(DataGridViewRow row in dgv.Rows){
                string line = "";
                foreach(DataGridViewCell cell in row.Cells){
                    if(cell.Value == null){
                        cell.Value = 0;
                    }
                    line += cell.Value + ",";
                }
                
                setting +=  line.Substring(0, line.Length - 1);
                setting += "\r\n";
               // MessageBox.Show(setting + "NNN");
            }
            
            return setting;
        }

        /// <summary>
        /// 获取当前网络信息设置
        /// </summary>
        private string getCurrSettingNet()
        {
            string setting = "";

            setting += tbxIp.Text;
            setting += "\r\n";
            setting += tbxPort.Text;
            return setting;
        }

        /// <summary>
        /// 获取之前保存的设置
        /// </summary>
        private void getSavedSetting()
        {
            string[,] setting = new string[12,15];
            string enter = "\r\n";
            String currPath = Environment.CurrentDirectory;
            String fileName = "L" + page + ".txt";
            FileInfo config = new FileInfo(currPath + "/" + fileName);
            if (!config.Exists)
            {
                config.Create().Close();
            }
            else
            {
                StreamReader reader = new StreamReader(config.FullName);
                string str = reader.ReadToEnd().Trim().Replace(enter, "E");
                reader.Dispose();
                reader.Close();
                //MessageBox.Show("trim长度：" + str.Length);
                if (str.Length >= SETTING_CONTENT_LENGTH)
                {
                    try
                    {
                        string[] lines = str.Split('E');

                        if (lines != null)
                        {
                            int i = 0;
                            foreach (string line in lines)
                            {
                                string[] values = line.Split(',');
                                int j = 0;
                                foreach(string value in values){
                                    dgv.Rows[i].Cells[j].Value = value;
                                    setCellBgColor(dgv.Rows[i].Cells[j]);
                                    j ++;
                                }
                                i ++;
                            }
                        }

                    }
                    catch(Exception e)
                    {
                        //tbxMsg.Text = "加载配置文件出错：\r\n" + e.Message;
                    }
                }
                else
                {
                    tbxMsg.Text = "配置文件格式有误";
                }
               
            }


        }


        /// <summary>
        /// 获取之前保存的网络信息设置
        /// </summary>
        private void getSavedSettingNet()
        {
         
            string enter = "\r\n";
            String currPath = Environment.CurrentDirectory;
            String fileName = "net.txt";
            FileInfo configNet = new FileInfo(currPath + "/" + fileName);
            if (!configNet.Exists)
            {
                configNet.Create().Close();
            }
            else
            {
                StreamReader reader = new StreamReader(configNet.FullName);
                tbxIp.Text = reader.ReadLine();
                tbxPort.Text = reader.ReadLine();
                reader.Dispose();
                reader.Close();
            }

        }

        /// <summary>
        /// 设置单元格背景色
        /// </summary>
        private void setCellBgColor(DataGridViewCell cell)
        {
            string value = cell.Value.ToString();
            if(value.Equals(WHITE + "")){
                cell.Style.BackColor = Color.White;
                cell.Style.ForeColor = Color.White;
            }else if(value.Equals(RED + "")){
                cell.Style.BackColor = Color.Red;
                cell.Style.ForeColor = Color.Red;
            }else if(value.Equals(GREEN + "")){
                cell.Style.BackColor = Color.Green;
                cell.Style.ForeColor = Color.Green;
            }else if(value.Equals(BLUE + "")){
                cell.Style.BackColor = Color.Blue;
                cell.Style.ForeColor = Color.Blue;
            }else if(value.Equals(YELLOW + "")){
                cell.Style.BackColor = Color.Yellow;
                cell.Style.ForeColor = Color.Yellow;
            }else if(value.Equals(ORANGE + "")){
                cell.Style.BackColor = Color.Orange;
                cell.Style.ForeColor = Color.Orange;
            }
           
        }

        /// <summary>
        /// 清空界面
        /// </summary>
        private void clear()
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                     cell.Value = 0;
                     cell.Style.ForeColor = Color.White;
                     cell.Style.BackColor = Color.White;
                }
            }
        }

        /// <summary>
        /// ???
        /// </summary>
        /// <param name="crc_array"></param>
        /// <returns></returns>
        public uint GetCRC(Byte[] crc_array)
        {
            uint i, j;
            uint modbus_crc;

            modbus_crc = 0xffff;

            for (i = 0; i < crc_array.Length - 2; i++)
            {
                modbus_crc = (modbus_crc & 0xFF00) | ((modbus_crc & 0x00FF) ^ crc_array[i]);
                for (j = 1; j <= 8; j++)
                {
                    if ((modbus_crc & 0x01) == 1)
                    {
                        modbus_crc = (modbus_crc >> 1);
                        modbus_crc ^= 0XA001;
                    }
                    else
                    {
                        modbus_crc = (modbus_crc >> 1);
                    }
                }
            }
            //MessageBox.Show(modbus_crc.ToString("x4").ToUpper() + "  Method 2");
            return modbus_crc;
        }
        /// <summary>
        /// Good !
        /// </summary>
        public byte[] CRC16MODBUS(byte[] dataBuff, byte dataLen)
        {
            byte CRC16High = 0;
            byte CRC16Low = 0;

            int CRCResult = 0xFFFF;
            for (int i = 0; i < dataLen; i++)
            {
                CRCResult = CRCResult ^ dataBuff[i];
                for (int j = 0; j < 8; j++)
                {
                    if ((CRCResult & 1) == 1)
                        CRCResult = (CRCResult >> 1) ^ 0xA001;
                    else
                        CRCResult >>= 1;
                }
            }
            CRC16High = Convert.ToByte(CRCResult & 0xff);
            CRC16Low = Convert.ToByte(CRCResult >> 8);

            byte[] result = new byte[2];
            result[0] = CRC16Low;
            result[1] = CRC16High;
            return result;
         
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
           
        }

        private void cbxPage_SelectedIndexChanged(object sender, EventArgs e)
        {
            page = cbxPage.SelectedIndex + 1;
            clear();           
            getSavedSetting();
        }

       

        
        /*
        private void label1_Click(object sender, EventArgs e)
        {
            byte[] bytes = {0xFF, 0x5A, 0x01, 0x5B};
            byte[] a = CRC16MODBUS(bytes, (byte)bytes.Length);
            lblMsg.Text = a[0].ToString("x2").ToUpper() + "  " + a[1].ToString("x2").ToUpper();
        }

         */
    }

    class MyMessage
    {
        string msg;

        public MyMessage(string str){
            msg = str;
        }

        public void Test(object para)
        {
            FormMain form = (FormMain)para;
            form.UpdateText(msg);
        }
    }

}
