using Panuon.UI.Silver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.IO;
using System.Windows.Threading;
using System.Security.Permissions;
using MySql.Data.MySqlClient;
using System.Security.Cryptography;

namespace OneForAll
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : WindowX
    {
        SerialPort SerialComPort1;
        #region 多线程定义
        Thread t_control = null;
        ThreadStart t_controlstart = null;
        Thread t_power = null;
        ThreadStart t_powerstart = null;
        Thread t_search = null;
        ThreadStart t_searchstart = null;
        Thread t_scan = null;
        ThreadStart t_scanstart = null;
        Thread t_inicfg = null;
        ThreadStart t_inicfgstart = null;
        Thread t_upgrade = null;
        ThreadStart t_upgradestart = null;
        Thread t_onekey = null;
        ThreadStart t_onekeystart = null;
        #endregion
        #region 画刷资源
        SolidColorBrush GreenBrush;
        SolidColorBrush RedBrush;
        SolidColorBrush DimGrayBrush;
        SolidColorBrush ButtonGrayBrush;
        SolidColorBrush ButtonBlueBrush;
        SolidColorBrush ButtonLightGrayBrush;
        SolidColorBrush TextWhiteBrush;
        SolidColorBrush TextBlackBrush;
        SolidColorBrush TextBlockWhiteBrush;
        SolidColorBrush TextBlockRedBrush;
        SolidColorBrush TextBlockGreenBrush;
        SolidColorBrush TabButtonGrayBrush;
        SolidColorBrush TabButtonLightGrayBrush;
        #endregion
        #region Buffer
        byte[] send1_buffer = new byte[400];
        byte[] read1_buffer = new byte[400];
        #endregion
        OpenFileDialog upgradefile = new OpenFileDialog();
        string modifypath;
        int buttonstate = 0;

        private MySqlConnection connection; 
        private string server;
        private string database;
        private string uid;
       // private string password;



        public MainWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
            if (INI.ReadIni("CONF", "language", AppDomain.CurrentDomain.BaseDirectory + "ONEFORALL.ini") == "English")
            {
 //               n_about.Text = "About";
  //              LanguageSelect.SelectedIndex = 0;
            }
            else if (INI.ReadIni("CONF", "language", AppDomain.CurrentDomain.BaseDirectory + "ONEFORALL.ini") == "简体中文")
            {
  //              n_about.Text = "关于";
  //              LanguageSelect.SelectedIndex = 1;
            }
            SerialComPort1 = new SerialPort();
            #region 画刷初始化
            GreenBrush = new SolidColorBrush(Colors.Green);
            RedBrush = new SolidColorBrush(Colors.Red);
            DimGrayBrush = new SolidColorBrush(Color.FromRgb(105, 105, 105));
            ButtonBlueBrush = new SolidColorBrush(Color.FromRgb(0, 122, 255));
            ButtonGrayBrush = new SolidColorBrush(Color.FromRgb(193, 193, 193));
            ButtonLightGrayBrush = new SolidColorBrush(Color.FromRgb(227, 227, 227));
            TextBlackBrush = new SolidColorBrush(Colors.Black);
            TextWhiteBrush = new SolidColorBrush(Colors.White);
            TextBlockWhiteBrush = new SolidColorBrush(Colors.White);
            TextBlockRedBrush = new SolidColorBrush(Colors.Red);
            TextBlockGreenBrush = new SolidColorBrush(Color.FromRgb(6, 176, 37));
            TabButtonGrayBrush = new SolidColorBrush(Color.FromRgb(221, 221, 221));
            TabButtonLightGrayBrush = new SolidColorBrush(Color.FromRgb(240, 240, 240));
            #endregion

            String connetStr = "server=rm-uf6lp2t237sr3af45to.mysql.rds.aliyuncs.com;" +
                               "user=tangtao_19891124;" +
                               "password=pP19891124!; " +
                               "database=tangtao;";
            // server=127.0.0.1/localhost 代表本机，端口号port默认是3306可以不写
            MySqlConnection conn = new MySqlConnection(connetStr);
            
            try
            {
                conn.Open();                        //打开通道，建立连接，可能出现异常,使用try catch语句
                Console.WriteLine("已经建立连接");
                                                    //在这里使用代码对数据库进行增删查改
            }
            catch (MySqlException ex)               //执行出错后的代码
            {
                Console.WriteLine(ex.Message);
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Cannot connect to server.  Contact administrator");
                        break;

                    case 1045:
                        MessageBox.Show("Invalid username/password, please try again");
                        break;
                }
            }
/*
            //事务
            //数据库事务( transaction)是访问并可能操作各种数据项的一个数据库操作序列，这些操作要么全部执行,
            //要么全部不执行，是一个不可分割的工作单位。事务由事务开始与事务结束之间执行的全部数据库操作组成。
            MySqlTransaction transaction = conn.BeginTransaction();//事务必须在try外面赋值不然catch里的transaction会报错:未赋值
            try
            {
                string date = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day;
                string sql1 = "insert into tangtao(username,password,registerdate) values('啊宽','123','" + date + "')";
                MySqlCommand cmd1 = new MySqlCommand(sql1, conn);
                cmd1.ExecuteNonQuery();

                string sql2 = "insert into tangtao(username,password,registerdate) values('啊宽2','123','" + date + "')";
                MySqlCommand cmd2 = new MySqlCommand(sql2, conn);
                cmd2.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                transaction.Rollback();//事务ExecuteNonQuery()执行失败报错，username被设置unique
                conn.Close();
            }
            finally
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                {
                    transaction.Commit();//事务要么回滚要么提交，即Rollback()与Commit()只能执行一个
                    conn.Close();
                }
            }
            int test = 0;
*/
            /*
                        //查询所有，从第一条到最后一条
                        string sql = "select * from tangtao";
                        MySqlCommand cmd = new MySqlCommand(sql, conn);
                        MySqlDataReader reader = cmd.ExecuteReader();//执行ExecuteReader()返回一个MySqlDataReader对象
                        while (reader.Read())//初始索引是-1，执行读取下一行数据，返回值是bool
                        {
                            //Console.WriteLine(reader[0].ToString() + reader[1].ToString() + reader[2].ToString());
                            //Console.WriteLine(reader.GetInt32(0)+reader.GetString(1)+reader.GetString(2));
                            Console.WriteLine(reader.GetString("username") + reader.GetString("password"));//"userid"是数据库对应的列名，推荐这种方式

                        }
                        int test = 0;
            */
            /*
                        //查询是否有数据
                        string sql = "select count(*) from tangtao";
                        MySqlCommand cmd = new MySqlCommand(sql, conn);
                        Object result = cmd.ExecuteScalar();//执行查询，并返回查询结果集中第一行的第一列。所有其他的列和行将被忽略。select语句无记录返回时，ExecuteScalar()返回NULL值
                        int count;
                        if (result != null)
                        {
                            count = int.Parse(result.ToString());
                        }
                        int test = 0;
            */

            /*            
                        //新建数据库
                        MySqlCommand cmd = new MySqlCommand("CREATE DATABASE tangtao1;", conn);
                        int result = cmd.ExecuteNonQuery();
            */
            /*
                        //新建表
                        //CREATE TABLE Test (Field1 VarChar(50), Field2 Integer)，这句话会建立一个名字Field1类型varchar长度50和一个名字Field2类型int
                        //CREATE TABLE Test (Field1 VarChar(50) primary key, Field2 Integer)，这句话会建立一个名字Field1类型varchar长度50有主键和一个名字Field2类型int
                        string createStatement = "CREATE TABLE Test (Field1 VarChar(50) primary key, Field2 Integer)";
                        MySqlCommand cmd = new MySqlCommand(createStatement, conn);
                        int result = cmd.ExecuteNonQuery(); //测试下来建表成功返回0
            */
            /*
                        //增加行
                        string alterStatement = "ALTER TABLE Test ADD Field3 Boolean";
                        MySqlCommand cmd = new MySqlCommand(alterStatement, conn);
                        int result = 1;
                        result = cmd.ExecuteNonQuery(); //测试下来增加行成功返回0
            */
            /*
                                    //MD5类是抽象类
                                    MD5 md5 = MD5.Create();
                                    //需要将字符串转成字节数组
                                    byte[] buffer = Encoding.Default.GetBytes("1234");
                                    //加密后是一个字节类型的数组，这里要注意编码UTF8/Unicode等的选择，这里的1234是密码
                                    byte[] md5buffer = md5.ComputeHash(buffer);
                                    string str = null;
                                    // 通过使用循环，将字节类型的数组转换为字符串，此字符串是常规字符格式化所得
                                    foreach (byte b in md5buffer)
                                    {
                                        //得到的字符串使用十六进制类型格式。格式后的字符是小写的字母，如果使用大写（X）则格式后的字符是大写字符 
                                        //但是在和对方测试过程中，发现我这边的MD5加密编码，经常出现少一位或几位的问题；
                                        //后来分析发现是 字符串格式符的问题， X 表示大写， x 表示小写， 
                                        //X2和x2表示不省略首位为0的十六进制数字；
                                        str += b.ToString("x2");
                                    }
            */


            /*
                        //我们知道如果插入数据时如果主键相同 或者有唯一索引之类的列数据相同 如果使用 insert into 插入会报错。
                        //使用 insert ignore into 如果遇到错误会忽略这个错误 ,然后返回 并没有对数据库进行操作，换句话说就是数据还是原来的数据
                        //没有错误就和insert into 一样
                        //string sql = "insert ignore into tangtao(username,password,registerdate) values('啊宽1','"+ str +"','" + DateTime.Now + "')";//插入
                        //string sql = "insert ignore into tangtao(username,password,registerdate) values('啊宽1','123','" + DateTime.Now + "')";//插入
                        //string sql = "delete from tangtao where username='啊宽1'";//删除username为啊宽1的这一行
                        //string sql = "update tangtao set username='啊宽1',password='1234' where username='啊宽1'";//更改

                        //insert into tangtao(username, password) values('啊宽1', '1234')  ON DUPLICATE KEY UPDATE  password = '12345'
                        //这句话username 为主键  如果存在username为啊宽1 的数据 就更新这条数据 password列为12345 没有就插入这条数据'啊宽1', '1234'
                        string sql = "insert into tangtao(username, password) values('啊宽1', '1234')  ON DUPLICATE KEY UPDATE  password = '12345'";
                        MySqlCommand cmd = new MySqlCommand(sql, conn);
                        //执行成功result为1
                        int result = cmd.ExecuteNonQuery();//3.执行插入、删除、更改语句。执行成功返回受影响的数据的行数，返回1可做true判断。执行失败不返回任何数据，报错，下面代码都不执行
                        int i = 0;
            */
            /*
                        string sql = "select * from tangtao where username=@para1 and password=@para2";//在sql语句中定义parameter，然后再给parameter赋值
                        MySqlCommand cmd = new MySqlCommand(sql, conn);

                        cmd.Parameters.AddWithValue("para1", "啊宽1");  //username
                        cmd.Parameters.AddWithValue("para2", str);      //password

                        MySqlDataReader reader = cmd.ExecuteReader();
                        if (reader.Read())                              //如果用户名和密码正确则能查询到一条语句，即读取下一行返回true
                        {
                            int i = 0;
                        }
            */

        }



        public class INI
        {
            /// <summary>
            /// 写操作
            /// </summary>
            /// <param name="section">节</param>
            /// <param name="key">键</param>
            /// <param name="value">值</param>
            /// <param name="filePath">文件路径</param>
            /// <returns></returns>
            [DllImport("kernel32")]
            private static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

            /// <summary>
            /// 读操作
            /// </summary>
            /// <param name="section">节</param>
            /// <param name="key">键</param>
            /// <param name="def">未读取到的默认值</param>
            /// <param name="retVal">读取到的值</param>
            /// <param name="size">大小</param>
            /// <param name="filePath">路径</param>
            /// <returns></returns>
            [DllImport("kernel32")]
            private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

            /// <summary>
            /// 读ini文件
            /// </summary>
            /// <param name="section">节</param>
            /// <param name="key">键</param>
            /// <param name="defValue">未读取到值时的默认值</param>
            /// <param name="filePath">文件路径</param>
            /// <returns></returns>
            public static string ReadIni(string section, string key, string path)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(section, key, "", temp, 255, path);
                return temp.ToString();
            }

            /// <summary>
            /// 写入ini文件
            /// </summary>
            /// <param name="section">节</param>
            /// <param name="key">键</param>
            /// <param name="value">值</param>
            /// <param name="path">文件路径</param>
            /// <returns></returns>
            public static void WriteIni(string section, string key, string value, string path)
            {
                WritePrivateProfileString(section, key, value, path);
            }
            /// <summary>
            /// 删除节
            /// </summary>
            /// <param name="section">节</param>
            /// <param name="filePath"></param>
            /// <returns></returns>
            public static long DeleteSection(string section)
            {
                string IniFilePath = @"F:\Setting.ini";
                return WritePrivateProfileString(section, null, null, IniFilePath);
            }

            /// <summary>
            /// 删除键
            /// </summary>
            /// <param name="section">节</param>
            /// <param name="key">键</param>
            /// <param name="filePath">文件路径</param>
            /// <returns></returns>
            public static long DeleteKey(string section, string key)
            {
                string IniFilePath = @"F:\Setting.ini";
                return WritePrivateProfileString(section, key, null, IniFilePath);
            }
        }

        public void threadinit()
        {
            t_inicfgstart = new ThreadStart(th_inicfg);
            t_inicfg = new Thread(t_inicfgstart);
            t_powerstart = new ThreadStart(th_power);
            t_power = new Thread(t_powerstart);
            t_controlstart = new ThreadStart(th_control);
            t_control = new Thread(t_controlstart);
            t_searchstart = new ThreadStart(th_search);
            t_search = new Thread(t_searchstart);
            t_upgradestart = new ThreadStart(th_upgrade);
            t_upgrade = new Thread(t_upgradestart);
            t_scanstart = new ThreadStart(th_scan);
            t_scan = new Thread(t_scanstart);
            t_onekeystart = new ThreadStart(th_onekey);
            t_onekey = new Thread(t_onekeystart);
        }

        private void ScanPortList(bool select_first_item)
        {
            PortSelect.Items.Clear();
            string[] port_number = SerialPort.GetPortNames();
            if (port_number.Length >= 1)
            {
                foreach (string port in port_number)
                {
                    PortSelect.Items.Add(port);
                }
                if (select_first_item == true)
                {
                    PortSelect.SelectedIndex = 0;
                }
            }
        }

        private void PortButon_Click(object sender, RoutedEventArgs e)
        {
            if (SerialComPort1.IsOpen == true)
            {
                SerialComPort1.Close();
            }
            else
            {
                if (PortSelect.SelectedItem.ToString() == null)
                {
                    MessageBox.Show("Port open failed !", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    try
                    {
                        SerialComPort1.PortName = PortSelect.SelectedItem.ToString();
                        SerialComPort1.BaudRate = 9600;
                        SerialComPort1.Parity = Parity.None;
                        SerialComPort1.StopBits = StopBits.One;
                        SerialComPort1.Open();
                    }
                    catch
                    {
                        MessageBox.Show("Port open failed !", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            if (SerialComPort1.IsOpen == true)
            {
                PortSelect.IsEnabled = false;
                PortButon.Content = "Close";
                switch (MainTabControl.SelectedIndex)
                {
                    case 0:
                        t_inicfg = new Thread(t_inicfgstart);
                        t_inicfg.Start();
                        break;
                    case 1:
                        t_power = new Thread(t_powerstart);
                        t_power.Start();
                        break;
                    case 2:
                        t_control = new Thread(t_controlstart);
                        t_control.Start();
                        break;
                    case 3:
                        t_search = new Thread(t_searchstart);
                        t_search.Start();
                        break;
                    case 4:
                        t_upgrade = new Thread(t_upgradestart);
                        t_upgrade.Start();
                        break;
                    case 5:
                        t_scan = new Thread(t_scanstart);
                        t_scan.Start();
                        break;
                    case 6:
                        t_onekey = new Thread(t_onekeystart);
                        t_onekey.Start();
                        break;
                    case 7:

                        break;

                }
            }
            else
            {
                PortSelect.IsEnabled = true;
                PortButon.Content = "open";
                if ((t_control.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_control.Abort();
                    t_control.Join();
                }
                if ((t_scan.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_scan.Abort();
                    t_scan.Join();
                }
                if ((t_upgrade.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_upgrade.Abort();
                    t_upgrade.Join();
                }
                if ((t_search.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_search.Abort();
                    t_search.Join();
                }
                if ((t_inicfg.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_inicfg.Abort();
                    t_inicfg.Join();
                }
                if ((t_power.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_power.Abort();
                    t_power.Join();
                }
                if ((t_onekey.ThreadState & ThreadState.Unstarted) != 0) { }
                else
                {
                    t_onekey.Abort();
                    t_onekey.Join();
                }
                buttonstate = 0;
            }
        }

        public void GetCRC16CheckCode(int len, out byte crcL, out byte crcH, params byte[] data)
        {
            UInt16 wCrc = 0xFFFF;

            for (int i = 0; i < len; i++)
            {
                wCrc ^= data[i];
                for (int j = 0; j < 8; j++)
                {
                    if ((wCrc & 1) != 0)
                    {
                        wCrc >>= 1;
                        wCrc ^= 0xA001;
                    }
                    else
                    {
                        wCrc >>= 1;
                    }
                }
            }

            crcL = (byte)wCrc;
            crcH = (byte)(wCrc >> 8);
        }
        public bool DataCheck(byte[] data, int len)
        {
            byte crcL = 0;
            byte crcH = 0;

            GetCRC16CheckCode(len, out crcL, out crcH, data);

            if ((crcH == 0) && (crcL == 0)) return true;
            else return false;
        }

        private void th_inicfg()
        {

        }
        private void th_onekey()
        {
            MessageBoxResult dr;
            uint send_count = 0;
            byte[] data = new byte[4];
            while (true)
            {
                if (buttonstate == 20)//一键参数LORA 
                {
                    OpenFileDialog sfd = new OpenFileDialog();
                    sfd.Filter = "ini文件|*.ini";
                    sfd.Multiselect = false;

                    if (sfd.ShowDialog() == true)
                    {
                        dr = MessageBox.Show(sfd.FileName, "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                        if (dr == MessageBoxResult.Cancel)
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("打开失败");
                        return;
                    }



                    while (buttonstate == 20)
                    {
                        if (buttonstate == 20)
                        {
                            onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                            {
                                if (onekeyallPLcheckbox.IsChecked == true)
                                {
                                    if (INI.ReadIni("Control", "id", sfd.FileName) == "0")
                                    {
                                        MessageBox.Show("ID = 0 or 子阵烧写完毕! All Control Box Finished");
                                        buttonstate = 0;
                                        return;//不知道为啥没有跳出while，后人可以解答一下
                                    }
                                }
                            }));
                        }
                        if (buttonstate == 0) return;//为了上面跳出while

                        onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            onekeyIDlabel.Dispatcher.Invoke(new Action(delegate
                            {
                                if (onekeyallPLcheckbox.IsChecked == true)
                                {
                                    onekeyIDlabel.Content = "当前烧写ID：" + INI.ReadIni("Control", "id", sfd.FileName);
                                }
                                else
                                {
                                    onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        onekeyIDlabel.Content = "当前烧写ID：" + onekeyIDtextbox.Text;
                                    }));
                                }
                            }));
                        }));


                        dr = MessageBox.Show("请连接LORA调试模块，连接完毕点击OK", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                        if (dr == MessageBoxResult.Cancel)
                        {
                            buttonstate = 0;
                            return;
                        }
                        #region 调试模块写入出厂参数
                        if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                        {
                            send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09; send1_buffer[3] = 0x00; send1_buffer[4] = 0x00;
                            send1_buffer[5] = 0x00; send1_buffer[6] = 0x66; send1_buffer[7] = 0x00;
                            if (INI.ReadIni("LORA-ybt", "module", sfd.FileName) == "0")
                            {
                                send1_buffer[8] = 0x17;
                            }
                            else
                            {
                                send1_buffer[8] = 0x12;
                            }
                            send1_buffer[9] = 0x03;
                            send1_buffer[10] = 0x00; send1_buffer[11] = 0x00;

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 12);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                        {
                            send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00; send1_buffer[4] = 0xaf;
                            send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                            send1_buffer[8] = 0x04;
                            send1_buffer[9] = 0x00;

                            send1_buffer[13] = 0x07;
                            send1_buffer[14] = 0x00;
                            send1_buffer[15] = 0x09;
                            send1_buffer[16] = 0x00;
                            send1_buffer[17] = 0x00;
                            send1_buffer[18] = 0x00;
                            send1_buffer[19] = 0x07;

                            send1_buffer[21] = 0x0d;
                            send1_buffer[22] = 0x0a;
                            if (INI.ReadIni("LORA-jxyl", "module", sfd.FileName) == "0")
                            {
                                send1_buffer[10] = 0x6c;
                                send1_buffer[11] = 0x40;
                                send1_buffer[12] = 0x12;

                                send1_buffer[20] = 0x73;
                            }
                            else
                            {
                                send1_buffer[10] = 0xe4;
                                send1_buffer[11] = 0xc0;
                                send1_buffer[12] = 0x26;

                                send1_buffer[20] = 0x7f;
                            }

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 23);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        #endregion
                        #region 控制箱上电尝试连接
                        dr = MessageBox.Show("控制箱请上电，上电完毕点击OK", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                        if (dr == MessageBoxResult.Cancel)
                        {
                            buttonstate = 0;
                            return;
                        }
                        send1_buffer[0] = 0x01; send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x01; send1_buffer[6] = 0x84; send1_buffer[7] = 0x0a;
                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 8);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 7 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("默认参数连接控制箱失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        #endregion
                        #region write id
                        send1_buffer[0] = 0x00; send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 81;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x01; send1_buffer[6] = 0x02;
                        onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (onekeyallPLcheckbox.IsChecked == true)
                            {
                                send1_buffer[7] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "id", sfd.FileName)) >> 8);
                                send1_buffer[8] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "id", sfd.FileName)) & 0xff);
                            }
                            else
                            {
                                onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                {
                                    send1_buffer[7] = Convert.ToByte(Convert.ToUInt16(onekeyIDtextbox.Text) >> 8);
                                    send1_buffer[8] = Convert.ToByte(Convert.ToUInt16(onekeyIDtextbox.Text) & 0xff);
                                }));
                            }
                        }));




                        GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                        SerialComPort1.Write(send1_buffer, 0, 11);
                        Thread.Sleep(1000);
                        #endregion
                        #region write time
                        send1_buffer[0] = 0x00;
                        send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 0x17;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x06; send1_buffer[6] = 0x0c;
                        DateTime SystemTime = DateTime.Now;
                        send1_buffer[7] = (byte)(SystemTime.Year >> 8);
                        send1_buffer[8] = (byte)(SystemTime.Year & 0xFF);
                        send1_buffer[9] = (byte)(SystemTime.Month >> 8);
                        send1_buffer[10] = (byte)(SystemTime.Month & 0xFF);
                        send1_buffer[11] = (byte)(SystemTime.Day >> 8);
                        send1_buffer[12] = (byte)(SystemTime.Day & 0xFF);
                        send1_buffer[13] = (byte)(SystemTime.Hour >> 8);
                        send1_buffer[14] = (byte)(SystemTime.Hour & 0xFF);
                        send1_buffer[15] = (byte)(SystemTime.Minute >> 8);
                        send1_buffer[16] = (byte)(SystemTime.Minute & 0xFF);
                        send1_buffer[17] = (byte)(SystemTime.Second >> 8);
                        send1_buffer[18] = (byte)(SystemTime.Second & 0xFF);
                        GetCRC16CheckCode(19, out send1_buffer[19], out send1_buffer[20], send1_buffer);
                        SerialComPort1.Write(send1_buffer, 0, 21);
                        Thread.Sleep(1000);
                        #endregion
                        #region write parameters
                        onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (onekeyallPLcheckbox.IsChecked == true)
                            {
                                send1_buffer[0] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "id", sfd.FileName)) & 0xff);
                            }
                            else
                            {
                                onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                {
                                    send1_buffer[0] = Convert.ToByte(onekeyIDtextbox.Text);
                                }));
                            }
                        }));
                        send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 0x2B;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x47; send1_buffer[6] = 0x8E;

                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "maxmotorcur", sfd.FileName)));
                        send1_buffer[7] = data[3];
                        send1_buffer[8] = data[2];
                        send1_buffer[9] = data[1];
                        send1_buffer[10] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "longitude", sfd.FileName)));
                        send1_buffer[11] = data[3];
                        send1_buffer[12] = data[2];
                        send1_buffer[13] = data[1];
                        send1_buffer[14] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "latitude", sfd.FileName)));
                        send1_buffer[15] = data[3];
                        send1_buffer[16] = data[2];
                        send1_buffer[17] = data[1];
                        send1_buffer[18] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "timezone", sfd.FileName)));
                        send1_buffer[19] = data[3];
                        send1_buffer[20] = data[2];
                        send1_buffer[21] = data[1];
                        send1_buffer[22] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "accuracy", sfd.FileName)));
                        send1_buffer[23] = data[3];
                        send1_buffer[24] = data[2];
                        send1_buffer[25] = data[1];
                        send1_buffer[26] = data[0];

                        send1_buffer[83] = 0;
                        send1_buffer[84] = send1_buffer[0];

                        send1_buffer[89] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "CleanAngle", sfd.FileName)) >> 8);
                        send1_buffer[90] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "CleanAngle", sfd.FileName)) & 0xff);

                        if (INI.ReadIni("Control", "orientation", sfd.FileName) == "east")
                        {
                            send1_buffer[91] = 0;
                            send1_buffer[92] = 0;
                        }
                        else if (INI.ReadIni("Control", "orientation", sfd.FileName) == "west")
                        {
                            send1_buffer[91] = 0;
                            send1_buffer[92] = 1;
                        }
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "modulewidth", sfd.FileName)));
                        send1_buffer[99] = data[3];
                        send1_buffer[100] = data[2];
                        send1_buffer[101] = data[1];
                        send1_buffer[102] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "postspace", sfd.FileName)));
                        send1_buffer[103] = data[3];
                        send1_buffer[104] = data[2];
                        send1_buffer[105] = data[1];
                        send1_buffer[106] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "terrslope", sfd.FileName)));
                        send1_buffer[107] = data[3];
                        send1_buffer[108] = data[2];
                        send1_buffer[109] = data[1];
                        send1_buffer[110] = data[0];

                        send1_buffer[119] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindUp", sfd.FileName)) >> 8);
                        send1_buffer[120] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindUp", sfd.FileName)) & 0xff);
                        send1_buffer[121] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeUp", sfd.FileName)) >> 8);
                        send1_buffer[122] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeUp", sfd.FileName)) & 0xff);
                        send1_buffer[123] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindDown", sfd.FileName)) >> 8);
                        send1_buffer[124] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindDown", sfd.FileName)) & 0xff);
                        send1_buffer[125] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeDown", sfd.FileName)) >> 8);
                        send1_buffer[126] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeDown", sfd.FileName)) & 0xff);
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "ESoftLimit", sfd.FileName)));
                        send1_buffer[127] = data[3];
                        send1_buffer[128] = data[2];
                        send1_buffer[129] = data[1];
                        send1_buffer[130] = data[0];
                        data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "WSoftLimit", sfd.FileName)));
                        send1_buffer[131] = data[3];
                        send1_buffer[132] = data[2];
                        send1_buffer[133] = data[1];
                        send1_buffer[134] = data[0];

                        send1_buffer[143] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "NightDock", sfd.FileName)) >> 8);
                        send1_buffer[144] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "NightDock", sfd.FileName)) & 0xff);
                        GetCRC16CheckCode(149, out send1_buffer[149], out send1_buffer[150], send1_buffer);
                        SerialComPort1.Write(send1_buffer, 0, 151);
                        Thread.Sleep(1000);
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        #endregion
                        #region write remote LORA
                        if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                        {
                            send1_buffer[0] = 0xcf; send1_buffer[1] = 0xcf; send1_buffer[2] = 0xc0; send1_buffer[3] = 0x00; send1_buffer[4] = 0x09;
                            send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                            send1_buffer[6] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                            send1_buffer[7] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                            send1_buffer[8] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                            send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                            send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                            send1_buffer[11] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                            send1_buffer[12] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                            send1_buffer[13] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 14);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 14 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                        {
                            send1_buffer[0] = 0xea; send1_buffer[1] = 0xeb; send1_buffer[2] = 0x12; send1_buffer[3] = 0x45;
                            send1_buffer[4] = 0xec; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                            #region 波特率
                            send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                            #endregion
                            #region 校验
                            send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                            #endregion
                            #region freq
                            UInt64 var1;
                            var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                            var1 = (var1 * 1000000000) / 61035;
                            send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                            send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                            send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                            #endregion
                            #region 扩频因子
                            send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                            #endregion
                            #region 模式
                            send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                            #endregion
                            #region 扩频带宽
                            send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                            #endregion
                            #region Node ID
                            send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                            send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                            #endregion
                            #region Net ID
                            send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                            #endregion
                            #region power
                            send1_buffer[19] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "power", sfd.FileName));
                            #endregion
                            int i = 0;
                            send1_buffer[20] = 0;
                            for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                            send1_buffer[21] = 0x0d;
                            send1_buffer[22] = 0x0a;

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 23);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 21 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        #endregion
                        #region write local LORA
                        if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                        {
                            send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09;
                            send1_buffer[3] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                            send1_buffer[4] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                            send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                            send1_buffer[6] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                            send1_buffer[7] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                            send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                            send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                            send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                            send1_buffer[11] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 12);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                        {
                            send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                            send1_buffer[4] = 0xaf; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                            #region 波特率
                            send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                            #endregion
                            #region 校验
                            send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                            #endregion
                            #region freq
                            UInt64 var1;
                            var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                            var1 = (var1 * 1000000000) / 61035;
                            send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                            send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                            send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                            #endregion
                            #region 扩频因子
                            send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                            #endregion
                            #region 模式
                            send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                            #endregion
                            #region 扩频带宽
                            send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                            #endregion
                            #region Node ID
                            send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                            send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                            #endregion
                            #region Net ID
                            send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                            #endregion
                            #region power
                            send1_buffer[19] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName)));
                            #endregion
                            int i = 0;
                            send1_buffer[20] = 0;
                            for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                            send1_buffer[21] = 0x0d;
                            send1_buffer[22] = 0x0a;

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 23);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        #endregion
                        #region 再次连接，确定正常
                        onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (onekeyallPLcheckbox.IsChecked == true)
                            {
                                send1_buffer[0] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "id", sfd.FileName)) & 0xff);
                            }
                            else
                            {
                                onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                {
                                    send1_buffer[0] = Convert.ToByte(onekeyIDtextbox.Text);
                                }));
                            }
                        }));
                        send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x01;
                        GetCRC16CheckCode(6, out send1_buffer[6], out send1_buffer[7], send1_buffer);
                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 8);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 7 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("连接控制箱失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        #endregion
                        onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (onekeyallPLcheckbox.IsChecked == true)
                            {
                                buttonstate = 20;
                                #region 修改配置文件ID
                                int id_bak, id_add, id_max;
                                id_add = Convert.ToInt16(INI.ReadIni("Control", "number", sfd.FileName));
                                id_bak = Convert.ToInt16(INI.ReadIni("Control", "id", sfd.FileName));
                                id_max = Convert.ToInt16(INI.ReadIni("Control", "idmax", sfd.FileName));
                                if (INI.ReadIni("Control", "idmode", sfd.FileName) == "+")
                                {
                                    if ((id_bak + id_add) > id_max)
                                    {
                                        INI.WriteIni("Control", "id", "0", sfd.FileName);
                                    }
                                    else
                                    {
                                        INI.WriteIni("Control", "id", (id_bak + id_add).ToString(), sfd.FileName);
                                    }
                                }
                                else if (INI.ReadIni("Control", "idmode", sfd.FileName) == "-")
                                {
                                    if ((id_bak + id_add) <= 0)
                                    {
                                        INI.WriteIni("Control", "id", "0", sfd.FileName);
                                    }
                                    else
                                    {
                                        INI.WriteIni("Control", "id", (id_bak - id_add).ToString(), sfd.FileName);
                                    }
                                }
                                #endregion

                            }
                            else
                            {
                                buttonstate = 0;
                                #region 修改配置文件ID
                                int id_bak = 0, id_add, id_max;
                                id_add = Convert.ToInt16(INI.ReadIni("Control", "number", sfd.FileName));
                                onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                {
                                    id_bak = Convert.ToInt16(onekeyIDtextbox.Text);
                                }));
                                id_max = Convert.ToInt16(INI.ReadIni("Control", "idmax", sfd.FileName));
                                if (INI.ReadIni("Control", "idmode", sfd.FileName) == "+")
                                {
                                    if ((id_bak + id_add) > id_max)
                                    {
                                        INI.WriteIni("Control", "id", "0", sfd.FileName);
                                    }
                                    else
                                    {
                                        INI.WriteIni("Control", "id", (id_bak + id_add).ToString(), sfd.FileName);
                                    }
                                }
                                else if (INI.ReadIni("Control", "idmode", sfd.FileName) == "-")
                                {
                                    if ((id_bak + id_add) <= 0)
                                    {
                                        INI.WriteIni("Control", "id", "0", sfd.FileName);
                                    }
                                    else
                                    {
                                        INI.WriteIni("Control", "id", (id_bak - id_add).ToString(), sfd.FileName);
                                    }
                                }
                                #endregion

                            }
                        }));
                        MessageBox.Show("烧写成功");
                    }
                }
                else if (buttonstate == 21)//一键消缺
                {
                    buttonstate = 0;
                    if (MessageBox.Show("已输入ID ? , ID Entered ?", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (onekeyIDEDtextbox.Text == "" || onekeyIDEDtextbox.Text == "输入ID")
                        {
                            MessageBox.Show("ID error");
                            return;
                        }
                    }));

                    OpenFileDialog sfd = new OpenFileDialog();
                    sfd.Filter = "ini文件|*.ini";
                    sfd.Multiselect = false;

                    if (sfd.ShowDialog() == true)
                    {
                        if (MessageBox.Show(sfd.FileName, "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.Cancel)
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("打开失败");
                        return;
                    }

                    onekeyIDEDlabel.Dispatcher.Invoke(new Action(delegate
                    {
                        onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                        {
                            onekeyIDEDlabel.Content = "当前烧写ID：" + onekeyIDEDtextbox.Text;
                        }));
                    }));

                    dr = MessageBox.Show("控制箱拆盖，串口线连接箱内LORA模块，完毕点击OK，Remove the cover of the control box, connect the serial port line to Lora module in the box, and click OK",
                                         "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (dr == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    #region write local LORA
                    if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                    {
                        send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09;
                        send1_buffer[3] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                        send1_buffer[4] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                        send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                        send1_buffer[6] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                        send1_buffer[7] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                        send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                        send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                        send1_buffer[11] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 12);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                    {
                        send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0xaf; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                        #region 波特率
                        send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                        #endregion
                        #region 校验
                        send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                        #endregion
                        #region freq
                        UInt64 var1;
                        var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                        var1 = (var1 * 1000000000) / 61035;
                        send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                        send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                        send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                        #endregion
                        #region 扩频因子
                        send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                        #endregion
                        #region 模式
                        send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                        #endregion
                        #region 扩频带宽
                        send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                        #endregion
                        #region Node ID
                        send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                        send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                        #endregion
                        #region Net ID
                        send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                        #endregion
                        #region power
                        send1_buffer[19] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName)));
                        #endregion
                        int i = 0;
                        send1_buffer[20] = 0;
                        for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                        send1_buffer[21] = 0x0d;
                        send1_buffer[22] = 0x0a;

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 23);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    #endregion
                    dr = MessageBox.Show("串口线连接箱内控制板，完毕点击OK，connect the serial port line to Control board in the box, and click OK",
                                         "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (dr == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    #region write id
                    send1_buffer[0] = 0x00; send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 81;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x01; send1_buffer[6] = 0x02;
                    onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[7] = Convert.ToByte(Convert.ToUInt16(onekeyIDEDtextbox.Text) >> 8);
                        send1_buffer[8] = Convert.ToByte(Convert.ToUInt16(onekeyIDEDtextbox.Text) & 0xff);
                    }));
                    GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                    SerialComPort1.Write(send1_buffer, 0, 11);
                    Thread.Sleep(1000);
                    #endregion
                    #region write time
                    send1_buffer[0] = 0x00;
                    send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 0x17;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x06; send1_buffer[6] = 0x0c;
                    DateTime SystemTime = DateTime.Now;
                    send1_buffer[7] = (byte)(SystemTime.Year >> 8);
                    send1_buffer[8] = (byte)(SystemTime.Year & 0xFF);
                    send1_buffer[9] = (byte)(SystemTime.Month >> 8);
                    send1_buffer[10] = (byte)(SystemTime.Month & 0xFF);
                    send1_buffer[11] = (byte)(SystemTime.Day >> 8);
                    send1_buffer[12] = (byte)(SystemTime.Day & 0xFF);
                    send1_buffer[13] = (byte)(SystemTime.Hour >> 8);
                    send1_buffer[14] = (byte)(SystemTime.Hour & 0xFF);
                    send1_buffer[15] = (byte)(SystemTime.Minute >> 8);
                    send1_buffer[16] = (byte)(SystemTime.Minute & 0xFF);
                    send1_buffer[17] = (byte)(SystemTime.Second >> 8);
                    send1_buffer[18] = (byte)(SystemTime.Second & 0xFF);
                    GetCRC16CheckCode(19, out send1_buffer[19], out send1_buffer[20], send1_buffer);
                    SerialComPort1.Write(send1_buffer, 0, 21);
                    Thread.Sleep(1000);
                    #endregion
                    #region write parameters
                    onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[0] = Convert.ToByte(onekeyIDEDtextbox.Text);
                    }));

                    send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 0x2B;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x47; send1_buffer[6] = 0x8E;

                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "maxmotorcur", sfd.FileName)));
                    send1_buffer[7] = data[3];
                    send1_buffer[8] = data[2];
                    send1_buffer[9] = data[1];
                    send1_buffer[10] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "longitude", sfd.FileName)));
                    send1_buffer[11] = data[3];
                    send1_buffer[12] = data[2];
                    send1_buffer[13] = data[1];
                    send1_buffer[14] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "latitude", sfd.FileName)));
                    send1_buffer[15] = data[3];
                    send1_buffer[16] = data[2];
                    send1_buffer[17] = data[1];
                    send1_buffer[18] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "timezone", sfd.FileName)));
                    send1_buffer[19] = data[3];
                    send1_buffer[20] = data[2];
                    send1_buffer[21] = data[1];
                    send1_buffer[22] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "accuracy", sfd.FileName)));
                    send1_buffer[23] = data[3];
                    send1_buffer[24] = data[2];
                    send1_buffer[25] = data[1];
                    send1_buffer[26] = data[0];

                    send1_buffer[83] = 0;
                    send1_buffer[84] = send1_buffer[0];

                    send1_buffer[89] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "CleanAngle", sfd.FileName)) >> 8);
                    send1_buffer[90] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "CleanAngle", sfd.FileName)) & 0xff);

                    if (INI.ReadIni("Control", "orientation", sfd.FileName) == "east")
                    {
                        send1_buffer[91] = 0;
                        send1_buffer[92] = 0;
                    }
                    else if (INI.ReadIni("Control", "orientation", sfd.FileName) == "west")
                    {
                        send1_buffer[91] = 0;
                        send1_buffer[92] = 1;
                    }
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "modulewidth", sfd.FileName)));
                    send1_buffer[99] = data[3];
                    send1_buffer[100] = data[2];
                    send1_buffer[101] = data[1];
                    send1_buffer[102] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "postspace", sfd.FileName)));
                    send1_buffer[103] = data[3];
                    send1_buffer[104] = data[2];
                    send1_buffer[105] = data[1];
                    send1_buffer[106] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "terrslope", sfd.FileName)));
                    send1_buffer[107] = data[3];
                    send1_buffer[108] = data[2];
                    send1_buffer[109] = data[1];
                    send1_buffer[110] = data[0];

                    send1_buffer[119] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindUp", sfd.FileName)) >> 8);
                    send1_buffer[120] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindUp", sfd.FileName)) & 0xff);
                    send1_buffer[121] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeUp", sfd.FileName)) >> 8);
                    send1_buffer[122] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeUp", sfd.FileName)) & 0xff);
                    send1_buffer[123] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindDown", sfd.FileName)) >> 8);
                    send1_buffer[124] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProWindDown", sfd.FileName)) & 0xff);
                    send1_buffer[125] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeDown", sfd.FileName)) >> 8);
                    send1_buffer[126] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "ProTimeDown", sfd.FileName)) & 0xff);
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "ESoftLimit", sfd.FileName)));
                    send1_buffer[127] = data[3];
                    send1_buffer[128] = data[2];
                    send1_buffer[129] = data[1];
                    send1_buffer[130] = data[0];
                    data = BitConverter.GetBytes(Convert.ToSingle(INI.ReadIni("Control", "WSoftLimit", sfd.FileName)));
                    send1_buffer[131] = data[3];
                    send1_buffer[132] = data[2];
                    send1_buffer[133] = data[1];
                    send1_buffer[134] = data[0];

                    send1_buffer[143] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "NightDock", sfd.FileName)) >> 8);
                    send1_buffer[144] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Control", "NightDock", sfd.FileName)) & 0xff);
                    GetCRC16CheckCode(149, out send1_buffer[149], out send1_buffer[150], send1_buffer);
                    SerialComPort1.Write(send1_buffer, 0, 151);
                    Thread.Sleep(1000);
                    if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    #endregion
                    dr = MessageBox.Show("串口线连接LORA调试模块，复原控制箱（盖子不用装），connect the serial port line to LORA debug module, Restore the control box (no need to install the cover) and then click OK",
                                         "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (dr == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    #region write local LORA
                    if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                    {
                        send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09;
                        send1_buffer[3] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                        send1_buffer[4] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                        send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                        send1_buffer[6] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                        send1_buffer[7] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                        send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                        send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                        send1_buffer[11] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 12);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                    {
                        send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0xaf; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                        #region 波特率
                        send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                        #endregion
                        #region 校验
                        send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                        #endregion
                        #region freq
                        UInt64 var1;
                        var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                        var1 = (var1 * 1000000000) / 61035;
                        send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                        send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                        send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                        #endregion
                        #region 扩频因子
                        send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                        #endregion
                        #region 模式
                        send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                        #endregion
                        #region 扩频带宽
                        send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                        #endregion
                        #region Node ID
                        send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                        send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                        #endregion
                        #region Net ID
                        send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                        #endregion
                        #region power
                        send1_buffer[19] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName)));
                        #endregion
                        int i = 0;
                        send1_buffer[20] = 0;
                        for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                        send1_buffer[21] = 0x0d;
                        send1_buffer[22] = 0x0a;

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 23);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    #endregion
                    #region 再次连接，确定正常
                    onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[0] = Convert.ToByte(onekeyIDEDtextbox.Text);
                    }));

                    send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x01;
                    GetCRC16CheckCode(6, out send1_buffer[6], out send1_buffer[7], send1_buffer);
                    send_count = 0;
                    do
                    {
                        send_count++;
                        SerialComPort1.Write(send1_buffer, 0, 8);
                        Thread.Sleep(1000);
                    }
                    while (SerialComPort1.BytesToRead != 7 && send_count <= 3);
                    if (send_count == 4)
                    {
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        MessageBox.Show("控制箱消缺失败，Control box defect elimination failed");
                        buttonstate = 0;
                        return;
                    }
                    if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    #endregion
                    MessageBox.Show("控制箱消缺成功，Control box defect elimination succeed");
                }
                else if (buttonstate == 22)//power写参数、LORA
                {
                    OpenFileDialog sfd = new OpenFileDialog();
                    sfd.Filter = "ini文件|*.ini";
                    sfd.Multiselect = false;

                    if (sfd.ShowDialog() == true)
                    {
                        dr = MessageBox.Show(sfd.FileName, "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                        if (dr == MessageBoxResult.Cancel)
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("打开失败");
                        return;
                    }
                    if (buttonstate == 22)
                    {
                        P_onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (P_onekeyallPLcheckbox.IsChecked == true)
                            {
                                if (INI.ReadIni("Power", "id", sfd.FileName) == "0")
                                {
                                    MessageBox.Show("ID = 0 or All Box Finished!");
                                    return;
                                }
                            }
                        }));
                    }

                    while (buttonstate == 22)
                    {
                        P_onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            P_onekeyIDlabel.Dispatcher.Invoke(new Action(delegate
                            {
                                if (P_onekeyallPLcheckbox.IsChecked == true)
                                {
                                    P_onekeyIDlabel.Content = "当前烧写ID：" + INI.ReadIni("Power", "id", sfd.FileName);
                                }
                                else
                                {
                                    P_onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_onekeyIDlabel.Content = "当前烧写ID：" + P_onekeyIDtextbox.Text;
                                    }));
                                }
                            }));
                        }));

                        dr = MessageBox.Show("请连接LORA调试模块，连接完毕点击OK", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                        if (dr == MessageBoxResult.Cancel)
                        {
                            buttonstate = 0;
                            return;
                        }
                        #region 调试模块写入出厂参数
                        if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                        {
                            send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09; send1_buffer[3] = 0x00; send1_buffer[4] = 0x00;
                            send1_buffer[5] = 0x00; send1_buffer[6] = 0x66; send1_buffer[7] = 0x00;
                            if (INI.ReadIni("LORA-ybt", "module", sfd.FileName) == "0")
                            {
                                send1_buffer[8] = 0x17;
                            }
                            else
                            {
                                send1_buffer[8] = 0x12;
                            }
                            send1_buffer[9] = 0x03;
                            send1_buffer[10] = 0x00; send1_buffer[11] = 0x00;

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 12);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                        {
                            send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00; send1_buffer[4] = 0xaf;
                            send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                            send1_buffer[8] = 0x04;
                            send1_buffer[9] = 0x00;

                            send1_buffer[13] = 0x07;
                            send1_buffer[14] = 0x00;
                            send1_buffer[15] = 0x09;
                            send1_buffer[16] = 0x00;
                            send1_buffer[17] = 0x00;
                            send1_buffer[18] = 0x00;
                            send1_buffer[19] = 0x07;

                            send1_buffer[21] = 0x0d;
                            send1_buffer[22] = 0x0a;
                            if (INI.ReadIni("LORA-jxyl", "module", sfd.FileName) == "0")
                            {
                                send1_buffer[10] = 0x6c;
                                send1_buffer[11] = 0x40;
                                send1_buffer[12] = 0x12;

                                send1_buffer[20] = 0x73;
                            }
                            else
                            {
                                send1_buffer[10] = 0xe4;
                                send1_buffer[11] = 0xc0;
                                send1_buffer[12] = 0x26;

                                send1_buffer[20] = 0x7f;
                            }

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 23);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        #endregion
                        #region 电源箱上电尝试连接
                        dr = MessageBox.Show("电源箱请上电，控制箱请断电，上电完毕点击OK", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                        if (dr == MessageBoxResult.Cancel)
                        {
                            buttonstate = 0;
                            return;
                        }
                        send1_buffer[0] = 0x01; send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x01; send1_buffer[6] = 0x84; send1_buffer[7] = 0x0a;
                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 8);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 7 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("默认参数连接控制箱失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        #endregion
                        #region write id
                        send1_buffer[0] = 0x00; send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 81;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x01; send1_buffer[6] = 0x02;
                        P_onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (P_onekeyallPLcheckbox.IsChecked == true)
                            {
                                send1_buffer[7] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Power", "id", sfd.FileName)) >> 8);
                                send1_buffer[8] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Power", "id", sfd.FileName)) & 0xff);
                            }
                            else
                            {
                                P_onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                {
                                    send1_buffer[7] = Convert.ToByte(Convert.ToUInt16(P_onekeyIDtextbox.Text) >> 8);
                                    send1_buffer[8] = Convert.ToByte(Convert.ToUInt16(P_onekeyIDtextbox.Text) & 0xff);
                                }));
                            }
                        }));

                        GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                        SerialComPort1.Write(send1_buffer, 0, 11);
                        Thread.Sleep(1000);
                        #endregion
                        #region write remote LORA
                        if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                        {
                            send1_buffer[0] = 0xcf; send1_buffer[1] = 0xcf; send1_buffer[2] = 0xc0; send1_buffer[3] = 0x00; send1_buffer[4] = 0x09;
                            send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                            send1_buffer[6] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                            send1_buffer[7] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                            send1_buffer[8] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                            send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                            send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                            send1_buffer[11] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                            send1_buffer[12] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                            send1_buffer[13] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 14);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 14 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                        {
                            send1_buffer[0] = 0xea; send1_buffer[1] = 0xeb; send1_buffer[2] = 0x12; send1_buffer[3] = 0x45;
                            send1_buffer[4] = 0xec; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                            #region 波特率
                            send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                            #endregion
                            #region 校验
                            send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                            #endregion
                            #region freq
                            UInt64 var1;
                            var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                            var1 = (var1 * 1000000000) / 61035;
                            send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                            send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                            send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                            #endregion
                            #region 扩频因子
                            send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                            #endregion
                            #region 模式
                            send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                            #endregion
                            #region 扩频带宽
                            send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                            #endregion
                            #region Node ID
                            send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                            send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                            #endregion
                            #region Net ID
                            send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                            #endregion
                            #region power
                            send1_buffer[19] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "power", sfd.FileName));
                            #endregion
                            int i = 0;
                            send1_buffer[20] = 0;
                            for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                            send1_buffer[21] = 0x0d;
                            send1_buffer[22] = 0x0a;

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 23);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 21 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        #endregion
                        #region write local LORA
                        if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                        {
                            send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09;
                            send1_buffer[3] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                            send1_buffer[4] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                            send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                            send1_buffer[6] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                            send1_buffer[7] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                            send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                            send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                             ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                             (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                            send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                            send1_buffer[11] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 12);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                        {
                            send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                            send1_buffer[4] = 0xaf; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                            #region 波特率
                            send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                            #endregion
                            #region 校验
                            send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                            #endregion
                            #region freq
                            UInt64 var1;
                            var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                            var1 = (var1 * 1000000000) / 61035;
                            send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                            send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                            send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                            #endregion
                            #region 扩频因子
                            send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                            #endregion
                            #region 模式
                            send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                            #endregion
                            #region 扩频带宽
                            send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                            #endregion
                            #region Node ID
                            send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                            send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                            #endregion
                            #region Net ID
                            send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                            #endregion
                            #region power
                            send1_buffer[19] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName)));
                            #endregion
                            int i = 0;
                            send1_buffer[20] = 0;
                            for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                            send1_buffer[21] = 0x0d;
                            send1_buffer[22] = 0x0a;

                            send_count = 0;
                            do
                            {
                                send_count++;
                                SerialComPort1.Write(send1_buffer, 0, 23);
                                Thread.Sleep(1000);
                            }
                            while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                            if (send_count == 4)
                            {
                                if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                                MessageBox.Show("本地LORA写参数失败");
                                buttonstate = 0;
                                return;
                            }
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        }
                        #endregion
                        #region 再次连接，确定正常
                        P_onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (P_onekeyallPLcheckbox.IsChecked == true)
                            {
                                send1_buffer[0] = Convert.ToByte(Convert.ToUInt16(INI.ReadIni("Power", "id", sfd.FileName)) & 0xff);
                            }
                            else
                            {
                                P_onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                                {
                                    send1_buffer[0] = Convert.ToByte(P_onekeyIDtextbox.Text);
                                }));
                            }
                        }));
                        send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0x00; send1_buffer[5] = 0x01;
                        GetCRC16CheckCode(6, out send1_buffer[6], out send1_buffer[7], send1_buffer);
                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 8);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 7 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("连接控制箱失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        #endregion
                        P_onekeyallPLcheckbox.Dispatcher.Invoke(new Action(delegate
                        {
                            if (P_onekeyallPLcheckbox.IsChecked == true)
                            {
                                buttonstate = 20;
                                #region 修改配置文件ID
                                int id_bak, id_add, id_max;
                                id_add = Convert.ToInt16(INI.ReadIni("Power", "number", sfd.FileName));
                                id_bak = Convert.ToInt16(INI.ReadIni("Power", "id", sfd.FileName));
                                id_max = Convert.ToInt16(INI.ReadIni("Power", "idmax", sfd.FileName));
                                if (INI.ReadIni("Power", "idmode", sfd.FileName) == "+")
                                {
                                    if ((id_bak + id_add) > id_max)
                                    {
                                        INI.WriteIni("Power", "id", "0", sfd.FileName);
                                    }
                                    else
                                    {
                                        INI.WriteIni("Power", "id", (id_bak + id_add).ToString(), sfd.FileName);
                                    }
                                }
                                else if (INI.ReadIni("Power", "idmode", sfd.FileName) == "-")
                                {
                                    if ((id_bak - id_add) <= 0)
                                    {
                                        INI.WriteIni("Power", "id", "0", sfd.FileName);
                                    }
                                    else
                                    {
                                        INI.WriteIni("Power", "id", (id_bak - id_add).ToString(), sfd.FileName);
                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                buttonstate = 0;
                            }
                        }));
                        MessageBox.Show("烧写成功");
                    }
                }
                else if (buttonstate == 23)//POWER 消缺
                {
                    buttonstate = 0;
                    if (MessageBox.Show("已输入ID ? , ID Entered ?", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    P_onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (P_onekeyIDEDtextbox.Text == "" || P_onekeyIDEDtextbox.Text == "输入ID")
                        {
                            MessageBox.Show("ID error");
                            return;
                        }
                    }));

                    OpenFileDialog sfd = new OpenFileDialog();
                    sfd.Filter = "ini文件|*.ini";
                    sfd.Multiselect = false;

                    if (sfd.ShowDialog() == true)
                    {
                        if (MessageBox.Show(sfd.FileName, "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.Cancel)
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("打开失败");
                        return;
                    }

                    P_onekeyIDEDlabel.Dispatcher.Invoke(new Action(delegate
                    {
                        P_onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                        {
                            P_onekeyIDEDlabel.Content = "当前烧写ID：" + P_onekeyIDEDtextbox.Text;
                        }));
                    }));

                    dr = MessageBox.Show("打开电源箱，串口线连接箱内LORA模块，完毕点击OK，Open the power box, connect the serial port line to Lora module in the box, and click OK",
                                         "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (dr == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    #region write local LORA
                    if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                    {
                        send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09;
                        send1_buffer[3] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                        send1_buffer[4] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                        send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                        send1_buffer[6] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                        send1_buffer[7] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                        send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                        send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                        send1_buffer[11] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 12);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                    {
                        send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0xaf; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                        #region 波特率
                        send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                        #endregion
                        #region 校验
                        send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                        #endregion
                        #region freq
                        UInt64 var1;
                        var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                        var1 = (var1 * 1000000000) / 61035;
                        send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                        send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                        send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                        #endregion
                        #region 扩频因子
                        send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                        #endregion
                        #region 模式
                        send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                        #endregion
                        #region 扩频带宽
                        send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                        #endregion
                        #region Node ID
                        send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                        send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                        #endregion
                        #region Net ID
                        send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                        #endregion
                        #region power
                        send1_buffer[19] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName)));
                        #endregion
                        int i = 0;
                        send1_buffer[20] = 0;
                        for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                        send1_buffer[21] = 0x0d;
                        send1_buffer[22] = 0x0a;

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 23);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    #endregion
                    dr = MessageBox.Show("串口线连接箱内控制板，完毕点击OK，connect the serial port line to Power board in the box, and click OK",
                                         "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (dr == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    #region write id
                    send1_buffer[0] = 0x00; send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 81;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x01; send1_buffer[6] = 0x02;
                    P_onekeyIDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[7] = Convert.ToByte(Convert.ToUInt16(P_onekeyIDEDtextbox.Text) >> 8);
                        send1_buffer[8] = Convert.ToByte(Convert.ToUInt16(P_onekeyIDEDtextbox.Text) & 0xff);
                    }));
                    GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                    SerialComPort1.Write(send1_buffer, 0, 11);
                    Thread.Sleep(1000);
                    #endregion
                    dr = MessageBox.Show("串口线连接LORA调试模块，复原电源箱（盖子不用装），connect the serial port line to LORA debug module, Restore the power box (no need to install the cover) and then click OK",
                                         "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (dr == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    #region write local LORA
                    if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                    {
                        send1_buffer[0] = 0xc0; send1_buffer[1] = 0x00; send1_buffer[2] = 0x09;
                        send1_buffer[3] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) >> 8);
                        send1_buffer[4] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName))) & 0xff);
                        send1_buffer[5] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "netid", sfd.FileName))));
                        send1_buffer[6] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName))));
                        send1_buffer[7] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName))) << 5) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName))));
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName))));
                        send1_buffer[9] = Convert.ToByte(((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName))) << 7) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName))) << 6) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName))) << 5) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName))) << 4) |
                                                         ((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName))) << 3) |
                                                         (Convert.ToUInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName))));
                        send1_buffer[10] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) >> 8);
                        send1_buffer[11] = Convert.ToByte((Convert.ToUInt16(INI.ReadIni("LORA-ybt", "key", sfd.FileName))) & 0xff);

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 12);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 12 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                    {
                        send1_buffer[0] = 0xaf; send1_buffer[1] = 0xaf; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                        send1_buffer[4] = 0xaf; send1_buffer[5] = 0x80; send1_buffer[6] = 0x01; send1_buffer[7] = 0x0c;
                        #region 波特率
                        send1_buffer[8] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName)));//波特率
                        #endregion
                        #region 校验
                        send1_buffer[9] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));//校验
                        #endregion
                        #region freq
                        UInt64 var1;
                        var1 = Convert.ToUInt64(INI.ReadIni("LORA-jxyl", "freq", sfd.FileName));//FREQ
                        var1 = (var1 * 1000000000) / 61035;
                        send1_buffer[10] = Convert.ToByte((var1 >> 16) & 0xff);
                        send1_buffer[11] = Convert.ToByte((var1 >> 8) & 0xff);
                        send1_buffer[12] = Convert.ToByte(var1 & 0xff);
                        #endregion
                        #region 扩频因子
                        send1_buffer[13] = Convert.ToByte(7 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName)));
                        #endregion
                        #region 模式
                        send1_buffer[14] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                        #endregion
                        #region 扩频带宽
                        send1_buffer[15] = Convert.ToByte(6 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName)));
                        #endregion
                        #region Node ID
                        send1_buffer[16] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName));
                        send1_buffer[17] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName));
                        #endregion
                        #region Net ID
                        send1_buffer[18] = Convert.ToByte(INI.ReadIni("LORA-jxyl", "netid", sfd.FileName));
                        #endregion
                        #region power
                        send1_buffer[19] = Convert.ToByte(1 + Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName)));
                        #endregion
                        int i = 0;
                        send1_buffer[20] = 0;
                        for (i = 0; i < 20; i++) send1_buffer[20] += send1_buffer[i];
                        send1_buffer[21] = 0x0d;
                        send1_buffer[22] = 0x0a;

                        send_count = 0;
                        do
                        {
                            send_count++;
                            SerialComPort1.Write(send1_buffer, 0, 23);
                            Thread.Sleep(1000);
                        }
                        while (SerialComPort1.BytesToRead != 23 && send_count <= 3);
                        if (send_count == 4)
                        {
                            if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                            MessageBox.Show("本地LORA写参数失败");
                            buttonstate = 0;
                            return;
                        }
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    }
                    #endregion
                    #region 再次连接，确定正常
                    P_onekeyIDEDtextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[0] = Convert.ToByte(P_onekeyIDEDtextbox.Text);
                    }));

                    send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x01;
                    GetCRC16CheckCode(6, out send1_buffer[6], out send1_buffer[7], send1_buffer);
                    send_count = 0;
                    do
                    {
                        send_count++;
                        SerialComPort1.Write(send1_buffer, 0, 8);
                        Thread.Sleep(1000);
                    }
                    while (SerialComPort1.BytesToRead != 7 && send_count <= 3);
                    if (send_count == 4)
                    {
                        if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                        MessageBox.Show("电源箱消缺失败，Power box defect elimination failed");
                        buttonstate = 0;
                        return;
                    }
                    if (SerialComPort1.BytesToRead != 0) SerialComPort1.Read(read1_buffer, 0, SerialComPort1.BytesToRead);
                    #endregion
                    MessageBox.Show("电源箱消缺成功，Power box defect elimination succeed");

                }

            }
        }


        private void th_power()
        {
            int bytes_to_read = 0, bytes_to_write = 0;
            byte[] data = new byte[4];
            while (true)
            {
                if (buttonstate == 0)
                {
                    PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                    }));
                    send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x72;
                    GetCRC16CheckCode(6, out send1_buffer[6], out send1_buffer[7], send1_buffer);
                    bytes_to_write = 8;
                }
                else if (buttonstate >= 3 && buttonstate <= 15)//mode
                {
                    P_BoardIDCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (P_BoardIDCheckBox.IsChecked == true)
                        {
                            send1_buffer[0] = 0x00;
                            P_BoardIDCheckBox.IsChecked = false;
                        }
                        else
                        {
                            PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                            {
                                send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                            }));
                        }
                    }));
                    send1_buffer[1] = 0x10;
                    send1_buffer[2] = 0;
                    send1_buffer[3] = 1;
                    send1_buffer[4] = 0;
                    send1_buffer[5] = 2;
                    send1_buffer[6] = 4;
                    send1_buffer[7] = 0;
                    send1_buffer[8] = 0;
                    send1_buffer[9] = 0;
                    switch (buttonstate)
                    {
                        case 6: send1_buffer[10] = 0x25; break;//lowerpower
                        case 14: send1_buffer[10] = 0xff; break;//c-reset
                        case 15: send1_buffer[10] = 0xfe; break;//c-sleep
                        default: break;
                    }
                    GetCRC16CheckCode(11, out send1_buffer[11], out send1_buffer[12], send1_buffer);
                    bytes_to_write = 13;
                }
                else if (buttonstate == 17)//change id
                {
                    P_BoardIDCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (P_BoardIDCheckBox.IsChecked == true)
                        {
                            send1_buffer[0] = 0x00;

                            P_BoardIDCheckBox.IsChecked = false;
                        }
                        else
                        {
                            PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                            {
                                send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                            }));
                        }
                    }));
                    PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        PortIDTextbox.Text = NewIDTextBox.Text;
                    }));

                    send1_buffer[1] = 0x10;
                    send1_buffer[2] = 0;
                    send1_buffer[3] = 0x51;
                    send1_buffer[4] = 0;
                    send1_buffer[5] = 1;
                    send1_buffer[6] = 2;
                    NewIDTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[7] = Convert.ToByte((Convert.ToUInt16(NewIDTextBox.Text)) >> 8);
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(NewIDTextBox.Text)) & 0xff);
                    }));
                    GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                    bytes_to_write = 11;
                }
                if (bytes_to_write != 0)
                {
                    SerialComPort1.Write(send1_buffer, 0, bytes_to_write);               
                    Thread.Sleep(1000);
                    if( buttonstate == 15 )
                    {
                        MessageBox.Show("已进入低功耗模式，点击确定后会唤醒 !");
                    }
                    buttonstate = 0;
                    bytes_to_read = SerialComPort1.BytesToRead;
                    if (bytes_to_read > 400)
                    {
                        while (bytes_to_read > 0)
                        {
                            if (bytes_to_read > 400)
                            {
                                SerialComPort1.Read(read1_buffer, 0, 400);
                            }
                            else
                            {
                                SerialComPort1.Read(read1_buffer, 0, 400);
                            }
                            bytes_to_read = SerialComPort1.BytesToRead;
                        }
                    }
                    else
                    {
                        SerialComPort1.Read(read1_buffer, 0, bytes_to_read);
                    }

                    if (bytes_to_read >= 5)
                    {
                        int id = 0;
                        PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                        {
                            id = Convert.ToByte(PortIDTextbox.Text);
                        }));

                        if ((read1_buffer[0] == id) && (DataCheck(read1_buffer, bytes_to_read) == true))
                        {
                            if (bytes_to_read == 233)
                            {
                                #region vision
                                P_HardwareTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_HardwareTextBlock.Text = "V" + read1_buffer[4].ToString();
                                }));
                                P_FirmwareTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_FirmwareTextBlock.Text = "V" + (read1_buffer[3] / 100.0).ToString();
                                }));
                                #endregion
                                #region workmode
                                Int32 workmode;
                                workmode = read1_buffer[1 * 2 + 3] << 24 | read1_buffer[1 * 2 + 3 + 1] << 16 | read1_buffer[2 * 2 + 3] << 8 | read1_buffer[2 * 2 + 3 + 1];

                                if (workmode == 0x25)
                                {
                                    P_LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_LowPowerButton.Background = ButtonBlueBrush;
                                        P_LowPowerButton.Foreground = TextWhiteBrush;
                                    }));

                                    P_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_ResetButton.Background = ButtonGrayBrush;
                                        P_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    P_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_SleepButton.Background = ButtonGrayBrush;
                                        P_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0xff)
                                {
                                    P_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_ResetButton.Background = ButtonGrayBrush;
                                        P_ResetButton.Foreground = TextBlackBrush;
                                    }));

                                    P_LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_LowPowerButton.Background = ButtonGrayBrush;
                                        P_LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    P_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_SleepButton.Background = ButtonGrayBrush;
                                        P_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0xfe)
                                {
                                    P_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_SleepButton.Background = ButtonBlueBrush;
                                        P_SleepButton.Foreground = TextWhiteBrush;
                                    }));

                                    P_LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_LowPowerButton.Background = ButtonGrayBrush;
                                        P_LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    P_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_ResetButton.Background = ButtonGrayBrush;
                                        P_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else
                                {
                                    P_LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_LowPowerButton.Background = ButtonGrayBrush;
                                        P_LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    P_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_ResetButton.Background = ButtonGrayBrush;
                                        P_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    P_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        P_SleepButton.Background = ButtonGrayBrush;
                                        P_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                #endregion
                                #region Alarm
                                data[0] = read1_buffer[12];
                                data[1] = read1_buffer[11];
                                data[2] = read1_buffer[10];
                                data[3] = read1_buffer[9];
                                uint AlarmMessage = BitConverter.ToUInt32(data, 0);
                                #region MotorOverCur
                                P_MotorOverCurIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & 1) > 0) P_MotorOverCurIndicator.Fill = RedBrush;
                                    else P_MotorOverCurIndicator.Fill = GreenBrush;
                                }));
                                #endregion                                

                                #region BatsocLow
                                P_BatSocLowIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 28)) > 0) P_BatSocLowIndicator.Fill = RedBrush;
                                    else P_BatSocLowIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region BatError
                                P_BatErrorIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 29)) > 0) P_BatErrorIndicator.Fill = RedBrush;
                                    else P_BatErrorIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region BatNoCom
                                P_BatNoComIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 31)) > 0) P_BatNoComIndicator.Fill = RedBrush;
                                    else P_BatNoComIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #endregion
                                P_SysTemperatureTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_SysTemperatureTextBlock.Text = Convert.ToUInt16(read1_buffer[81] << 8 | read1_buffer[82]).ToString() + "℃";
                                }));
                                P_SysCurTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_SysCurTextBlock.Text = (Convert.ToUInt16(read1_buffer[83] << 8 | read1_buffer[84]) / 10.0).ToString("F01") + "A";
                                }));
                                #region Power
                                P_PVBuckVolTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_PVBuckVolTextBlock.Text = (Convert.ToUInt16(read1_buffer[113] << 8 | read1_buffer[114]) / 10.0).ToString("F01") + "V";
                                }));
                                P_BatCurTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    data[0] = read1_buffer[59 * 2 + 3 + 1];
                                    data[1] = read1_buffer[59 * 2 + 3];
                                    P_BatCurTextBlock.Text = (BitConverter.ToInt16(data, 0) / 100.0).ToString("F02") + "A";
                                }));
                                P_BatVolSocTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_BatVolSocTextBlock.Text = (Convert.ToUInt16(read1_buffer[58 * 2 + 3] << 8 | read1_buffer[58 * 2 + 3 + 1]) / 10.0).ToString("F01") + "V/" + read1_buffer[57 * 2 + 3].ToString() + "%";
                                }));
                                P_BatTempTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    P_BatTempTextBlock.Text = ((sbyte)read1_buffer[57 * 2 + 4]).ToString() + "℃";
                                }));
                                #endregion
                                #region charger
                                ushort PowerState = Convert.ToUInt16(read1_buffer[56 * 2 + 3] << 8 | read1_buffer[56 * 2 + 3 + 1]);
                                P_HeatIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & 1) > 0) P_HeatIndicator.Fill = GreenBrush;
                                    else P_HeatIndicator.Fill = DimGrayBrush;
                                }));
                                P_PreChargeIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 1)) > 0) P_PreChargeIndicator.Fill = GreenBrush;
                                    else P_PreChargeIndicator.Fill = DimGrayBrush;
                                }));
                                P_CCCVIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 2)) > 0) P_CCCVIndicator.Fill = GreenBrush;
                                    else P_CCCVIndicator.Fill = DimGrayBrush;
                                }));
                                P_TerminationIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 3)) > 0) P_TerminationIndicator.Fill = GreenBrush;
                                    else P_TerminationIndicator.Fill = DimGrayBrush;
                                }));
                                P_MaxTimeIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 4)) > 0) P_MaxTimeIndicator.Fill = RedBrush;
                                    else P_MaxTimeIndicator.Fill = DimGrayBrush;
                                }));
                                P_BatMissingIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 5)) > 0) P_BatMissingIndicator.Fill = RedBrush;
                                    else P_BatMissingIndicator.Fill = DimGrayBrush;
                                }));
                                P_BatShortIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 6)) > 0) P_BatShortIndicator.Fill = RedBrush;
                                    else P_BatShortIndicator.Fill = DimGrayBrush;
                                }));
                                #endregion
                            }
                            else
                            {

                            }
                        }
                    }
                }
            }
        }

        private void th_control()
        {
            int bytes_to_read = 0, bytes_to_write = 0;
            byte[] data = new byte[4];
            while (true)
            {
                if (buttonstate == 0 || buttonstate == 1)
                {
                    PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                    }));
                    send1_buffer[1] = 0x03; send1_buffer[2] = 0x00; send1_buffer[3] = 0x00;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x72;
                    GetCRC16CheckCode(6, out send1_buffer[6], out send1_buffer[7], send1_buffer);
                    bytes_to_write = 8;
                }
                else if (buttonstate == 2)
                {
                    Array.Clear(send1_buffer, 0, send1_buffer.Length);
                    PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                    }));
                    send1_buffer[1] = 0x10; send1_buffer[2] = 0x00; send1_buffer[3] = 0x2B;
                    send1_buffer[4] = 0x00; send1_buffer[5] = 0x47; send1_buffer[6] = 0x8E;
                    MaxMotorCurTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(MaxMotorCurTextBox.Text));
                    }));
                    send1_buffer[7] = data[3];
                    send1_buffer[8] = data[2];
                    send1_buffer[9] = data[1];
                    send1_buffer[10] = data[0];

                    LongitudeTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(LongitudeTextBox.Text));
                    }));
                    send1_buffer[11] = data[3];
                    send1_buffer[12] = data[2];
                    send1_buffer[13] = data[1];
                    send1_buffer[14] = data[0];

                    LatitudeTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(LatitudeTextBox.Text));
                    }));
                    send1_buffer[15] = data[3];
                    send1_buffer[16] = data[2];
                    send1_buffer[17] = data[1];
                    send1_buffer[18] = data[0];

                    TimeZoneTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(TimeZoneTextBox.Text));
                    }));
                    send1_buffer[19] = data[3];
                    send1_buffer[20] = data[2];
                    send1_buffer[21] = data[1];
                    send1_buffer[22] = data[0];

                    AccuracyTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(AccuracyTextBox.Text));
                    }));
                    send1_buffer[23] = data[3];
                    send1_buffer[24] = data[2];
                    send1_buffer[25] = data[1];
                    send1_buffer[26] = data[0];

                    PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[83] = Convert.ToByte(Convert.ToUInt16(PortIDTextbox.Text) >> 8);
                        send1_buffer[84] = Convert.ToByte(Convert.ToUInt16(PortIDTextbox.Text) & 0xff);
                    }));

                    CleanAngleTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[89] = Convert.ToByte(Convert.ToUInt16(CleanAngleTextBox.Text) >> 8);
                        send1_buffer[90] = Convert.ToByte(Convert.ToUInt16(CleanAngleTextBox.Text) & 0xff);
                    }));

                    EastwardCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (EastwardCheckBox.IsChecked == true)
                        {
                            send1_buffer[91] = 0;
                            send1_buffer[92] = 1;
                        }
                        else
                        {
                            send1_buffer[91] = 0;
                            send1_buffer[92] = 0;
                        }
                    }));

                    ModuleWidthTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(ModuleWidthTextBox.Text));
                    }));
                    send1_buffer[99] = data[3];
                    send1_buffer[100] = data[2];
                    send1_buffer[101] = data[1];
                    send1_buffer[102] = data[0];

                    PostSpaceTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(PostSpaceTextBox.Text));
                    }));
                    send1_buffer[103] = data[3];
                    send1_buffer[104] = data[2];
                    send1_buffer[105] = data[1];
                    send1_buffer[106] = data[0];

                    TerrSlopeTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(TerrSlopeTextBox.Text));
                    }));
                    send1_buffer[107] = data[3];
                    send1_buffer[108] = data[2];
                    send1_buffer[109] = data[1];
                    send1_buffer[110] = data[0];

                    ProWindUpTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[119] = Convert.ToByte(Convert.ToUInt16(ProWindUpTextBox.Text) >> 8);
                        send1_buffer[120] = Convert.ToByte(Convert.ToUInt16(ProWindUpTextBox.Text) & 0xff);
                    }));

                    ProTimeUpTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[121] = Convert.ToByte(Convert.ToUInt16(ProTimeUpTextBox.Text) >> 8);
                        send1_buffer[122] = Convert.ToByte(Convert.ToUInt16(ProTimeUpTextBox.Text) & 0xff);
                    }));

                    ProWindDownTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[123] = Convert.ToByte(Convert.ToUInt16(ProWindDownTextBox.Text) >> 8);
                        send1_buffer[124] = Convert.ToByte(Convert.ToUInt16(ProWindDownTextBox.Text) & 0xff);
                    }));

                    ProTimeDownTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[125] = Convert.ToByte(Convert.ToUInt16(ProTimeDownTextBox.Text) >> 8);
                        send1_buffer[126] = Convert.ToByte(Convert.ToUInt16(ProTimeDownTextBox.Text) & 0xff);
                    }));

                    ESoftLimitTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(ESoftLimitTextBox.Text));
                    }));
                    send1_buffer[127] = data[3];
                    send1_buffer[128] = data[2];
                    send1_buffer[129] = data[1];
                    send1_buffer[130] = data[0];

                    WSoftLimitTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        data = BitConverter.GetBytes(Convert.ToSingle(WSoftLimitTextBox.Text));
                    }));
                    send1_buffer[131] = data[3];
                    send1_buffer[132] = data[2];
                    send1_buffer[133] = data[1];
                    send1_buffer[134] = data[0];

                    NightDockTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[143] = Convert.ToByte(Convert.ToUInt16(NightDockTextBox.Text) >> 8);
                        send1_buffer[144] = Convert.ToByte(Convert.ToUInt16(NightDockTextBox.Text) & 0xff);
                    }));

                    GetCRC16CheckCode(149, out send1_buffer[149], out send1_buffer[150], send1_buffer);
                    bytes_to_write = 151;
                }
                else if (buttonstate >= 3 && buttonstate <= 15)//mode
                {
                    C_BoardIDCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (C_BoardIDCheckBox.IsChecked == true)
                        {
                            send1_buffer[0] = 0x00;
                            C_BoardIDCheckBox.IsChecked = false;
                        }
                        else
                        {
                            PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                            {
                                send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                            }));
                        }
                    }));

                    send1_buffer[1] = 0x10;
                    send1_buffer[2] = 0;
                    send1_buffer[3] = 1;
                    send1_buffer[4] = 0;
                    send1_buffer[5] = 2;
                    send1_buffer[6] = 4;
                    send1_buffer[7] = 0;
                    send1_buffer[8] = 0;
                    send1_buffer[9] = 0;
                    switch (buttonstate)
                    {
                        case 3: send1_buffer[10] = 0x20; break;//auto
                        case 4: send1_buffer[10] = 0x28; break;//ai
                        case 5: send1_buffer[10] = 0x24; break;//wind
                        case 6: send1_buffer[10] = 0x25; break;//lowerpower
                        case 7: send1_buffer[10] = 0x22; break;//snow
                        case 8: send1_buffer[10] = 0x21; break;//rain
                        case 9: send1_buffer[10] = 0x18; break;//east
                        case 10: send1_buffer[10] = 0x14; break;//west
                        case 11: send1_buffer[10] = 0x60; break;//clean
                        case 12: send1_buffer[10] = 0x80; break;//cali
                        case 13: send1_buffer[10] = 0x40; break;//horizontal
                        case 14: send1_buffer[10] = 0xff; break;//c-reset
                        case 15: send1_buffer[10] = 0xfe; break;//c-sleep
                        default: break;
                    }
                    GetCRC16CheckCode(11, out send1_buffer[11], out send1_buffer[12], send1_buffer);
                    bytes_to_write = 13;
                }
                else if (buttonstate == 16)//ai angle
                {
                    C_BoardIDCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (C_BoardIDCheckBox.IsChecked == true)
                        {
                            send1_buffer[0] = 0x00;
                            C_BoardIDCheckBox.IsChecked = false;
                        }
                        else
                        {
                            PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                            {
                                send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                            }));
                        }
                    }));

                    send1_buffer[1] = 0x10;
                    send1_buffer[2] = 0;
                    send1_buffer[3] = 0x22;
                    send1_buffer[4] = 0;
                    send1_buffer[5] = 1;
                    send1_buffer[6] = 2;
                    AIAngleTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[7] = Convert.ToByte((Convert.ToUInt16(AIAngleTextBox.Text)) >> 8);
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(AIAngleTextBox.Text)) & 0xff);
                    }));
                    GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                    bytes_to_write = 11;
                }
                else if (buttonstate == 17)//change id
                {
                    C_BoardIDCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (C_BoardIDCheckBox.IsChecked == true)
                        {
                            send1_buffer[0] = 0x00;
                            C_BoardIDCheckBox.IsChecked = false;
                        }
                        else
                        {
                            PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                            {
                                send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                            }));
                        }
                    }));

                    PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                    {
                        PortIDTextbox.Text = NewIDTextBox.Text;
                    }));
                    send1_buffer[1] = 0x10;
                    send1_buffer[2] = 0;
                    send1_buffer[3] = 0x51;
                    send1_buffer[4] = 0;
                    send1_buffer[5] = 1;
                    send1_buffer[6] = 2;
                    NewIDTextBox.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[7] = Convert.ToByte((Convert.ToUInt16(NewIDTextBox.Text)) >> 8);
                        send1_buffer[8] = Convert.ToByte((Convert.ToUInt16(NewIDTextBox.Text)) & 0xff);
                    }));
                    GetCRC16CheckCode(9, out send1_buffer[9], out send1_buffer[10], send1_buffer);
                    bytes_to_write = 11;
                }
                else if (buttonstate == 18)//change time
                {
                    C_BoardIDCheckBox.Dispatcher.Invoke(new Action(delegate
                    {
                        if (C_BoardIDCheckBox.IsChecked == true)
                        {
                            send1_buffer[0] = 0x00;
                            C_BoardIDCheckBox.IsChecked = false;
                        }
                        else
                        {
                            PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                            {
                                send1_buffer[0] = Convert.ToByte(PortIDTextbox.Text);
                            }));
                        }
                    }));

                    send1_buffer[1] = 0x10;
                    send1_buffer[2] = 0;
                    send1_buffer[3] = 0x17;
                    send1_buffer[4] = 0;
                    send1_buffer[5] = 6;
                    send1_buffer[6] = 0x0c;
                    DateTime SystemTime = DateTime.Now;
                    send1_buffer[7] = (byte)(SystemTime.Year >> 8);
                    send1_buffer[8] = (byte)(SystemTime.Year & 0xFF);
                    send1_buffer[9] = (byte)(SystemTime.Month >> 8);
                    send1_buffer[10] = (byte)(SystemTime.Month & 0xFF);
                    send1_buffer[11] = (byte)(SystemTime.Day >> 8);
                    send1_buffer[12] = (byte)(SystemTime.Day & 0xFF);
                    send1_buffer[13] = (byte)(SystemTime.Hour >> 8);
                    send1_buffer[14] = (byte)(SystemTime.Hour & 0xFF);
                    send1_buffer[15] = (byte)(SystemTime.Minute >> 8);
                    send1_buffer[16] = (byte)(SystemTime.Minute & 0xFF);
                    send1_buffer[17] = (byte)(SystemTime.Second >> 8);
                    send1_buffer[18] = (byte)(SystemTime.Second & 0xFF);
                    GetCRC16CheckCode(19, out send1_buffer[19], out send1_buffer[20], send1_buffer);
                    bytes_to_write = 21;
                }
                if (bytes_to_write != 0)
                {
                    SerialComPort1.Write(send1_buffer, 0, bytes_to_write);
                    Thread.Sleep(1000);
                    if (buttonstate == 15)
                    {
                        MessageBox.Show("已进入低功耗模式，点击确定后会唤醒 !");
                    }
                    buttonstate = 0;
                    bytes_to_read = SerialComPort1.BytesToRead;
                    if( bytes_to_read >400 )
                    {
                        while(bytes_to_read>0)
                        {
                            if (bytes_to_read > 400)
                            {
                                SerialComPort1.Read(read1_buffer, 0, 400);
                            }
                            else
                            {
                                SerialComPort1.Read(read1_buffer, 0, 400);
                            }
                            bytes_to_read = SerialComPort1.BytesToRead;
                        }                      
                    }
                    else
                    {
                        SerialComPort1.Read(read1_buffer, 0, bytes_to_read);
                    }
                    if (bytes_to_read >= 5)
                    {
                        int id = 0;
                        PortIDTextbox.Dispatcher.Invoke(new Action(delegate
                        {
                            id = Convert.ToByte(PortIDTextbox.Text);
                        }));

                        if ((read1_buffer[0] == id) && (DataCheck(read1_buffer, bytes_to_read) == true))
                        {
                            if (bytes_to_read == 233)
                            {
                                #region vision
                                HardwareTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    HardwareTextBlock.Text = "V" + read1_buffer[4].ToString();
                                }));
                                FirmwareTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    FirmwareTextBlock.Text = "V" + (read1_buffer[3] / 100.0).ToString();
                                }));
                                #endregion
                                #region workmode
                                Int32 workmode;
                                workmode = read1_buffer[1 * 2 + 3] << 24 | read1_buffer[1 * 2 + 3 + 1] << 16 | read1_buffer[2 * 2 + 3] << 8 | read1_buffer[2 * 2 + 3 + 1];
                                if (workmode == 0x20)
                                {
                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonBlueBrush;
                                        AutoButton.Foreground = TextWhiteBrush;
                                    }));

                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x28)
                                {
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonBlueBrush;
                                        AIButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x24)
                                {
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonBlueBrush;
                                        WindButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x25)
                                {
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonBlueBrush;
                                        LowPowerButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x22)
                                {
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonBlueBrush;
                                        SnowButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x21)
                                {
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonBlueBrush;
                                        RainButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x18)
                                {
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonBlueBrush;
                                        EastButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x14)
                                {
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonBlueBrush;
                                        WestButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x60)
                                {
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonBlueBrush;
                                        CleanButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x80)
                                {
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0x40)
                                {
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonBlueBrush;
                                        HorizontalButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0xff)
                                {
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else if (workmode == 0xfe)
                                {
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonBlueBrush;
                                        C_SleepButton.Foreground = TextWhiteBrush;
                                    }));

                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                else
                                {
                                    AutoButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AutoButton.Background = ButtonGrayBrush;
                                        AutoButton.Foreground = TextBlackBrush;
                                    }));
                                    AIButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AIButton.Background = ButtonGrayBrush;
                                        AIButton.Foreground = TextBlackBrush;
                                    }));
                                    WindButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WindButton.Background = ButtonGrayBrush;
                                        WindButton.Foreground = TextBlackBrush;
                                    }));
                                    LowPowerButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LowPowerButton.Background = ButtonGrayBrush;
                                        LowPowerButton.Foreground = TextBlackBrush;
                                    }));
                                    SnowButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        SnowButton.Background = ButtonGrayBrush;
                                        SnowButton.Foreground = TextBlackBrush;
                                    }));
                                    RainButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        RainButton.Background = ButtonGrayBrush;
                                        RainButton.Foreground = TextBlackBrush;
                                    }));
                                    EastButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        EastButton.Background = ButtonGrayBrush;
                                        EastButton.Foreground = TextBlackBrush;
                                    }));
                                    WestButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WestButton.Background = ButtonGrayBrush;
                                        WestButton.Foreground = TextBlackBrush;
                                    }));
                                    CleanButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanButton.Background = ButtonGrayBrush;
                                        CleanButton.Foreground = TextBlackBrush;
                                    }));
                                    CaliButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CaliButton.Background = ButtonGrayBrush;
                                        CaliButton.Foreground = TextBlackBrush;
                                    }));
                                    HorizontalButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        HorizontalButton.Background = ButtonGrayBrush;
                                        HorizontalButton.Foreground = TextBlackBrush;
                                    }));
                                    C_ResetButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_ResetButton.Background = ButtonGrayBrush;
                                        C_ResetButton.Foreground = TextBlackBrush;
                                    }));
                                    C_SleepButton.Dispatcher.Invoke(new Action(delegate
                                    {
                                        C_SleepButton.Background = ButtonGrayBrush;
                                        C_SleepButton.Foreground = TextBlackBrush;
                                    }));
                                }
                                #endregion
                                #region Alarm
                                data[0] = read1_buffer[12];
                                data[1] = read1_buffer[11];
                                data[2] = read1_buffer[10];
                                data[3] = read1_buffer[9];
                                uint AlarmMessage = BitConverter.ToUInt32(data, 0);
                                #region MotorOverCur
                                MotorOverCurIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & 1) > 0) MotorOverCurIndicator.Fill = RedBrush;
                                    else MotorOverCurIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region MotorReversal
                                MotorReversalIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 1)) > 0) MotorReversalIndicator.Fill = RedBrush;
                                    else MotorReversalIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region ESoftlimit
                                ESoftLimitIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 2)) > 0) ESoftLimitIndicator.Fill = RedBrush;
                                    else ESoftLimitIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region WSoftlimit
                                WSoftLimitIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 3)) > 0) WSoftLimitIndicator.Fill = RedBrush;
                                    else WSoftLimitIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region AngleNoCom
                                AngleNoComIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 4)) > 0) AngleNoComIndicator.Fill = RedBrush;
                                    else AngleNoComIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region AngleNoChange
                                AngleNoChangeIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 5)) > 0) AngleNoChangeIndicator.Fill = RedBrush;
                                    else AngleNoChangeIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region RTCError
                                RTCErrorIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 24)) > 0) RTCErrorIndicator.Fill = RedBrush;
                                    else RTCErrorIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region emergency
                                EmergencyStopIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 26)) > 0) EmergencyStopIndicator.Fill = RedBrush;
                                    else EmergencyStopIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region BatsocLow
                                BatSocLowIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 28)) > 0) BatSocLowIndicator.Fill = RedBrush;
                                    else BatSocLowIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region BatError
                                BatErrorIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 29)) > 0) BatErrorIndicator.Fill = RedBrush;
                                    else BatErrorIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region TimeLost
                                TimeLostIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 30)) > 0) TimeLostIndicator.Fill = RedBrush;
                                    else TimeLostIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #region BatNoCom
                                BatNoComIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((AlarmMessage & (1 << 31)) > 0) BatNoComIndicator.Fill = RedBrush;
                                    else BatNoComIndicator.Fill = GreenBrush;
                                }));
                                #endregion

                                #endregion
                                #region angle data
                                data[0] = read1_buffer[20];
                                data[1] = read1_buffer[19];
                                data[2] = read1_buffer[18];
                                data[3] = read1_buffer[17];
                                ElevationTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    ElevationTextBlock.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                }));
                                data[0] = read1_buffer[24];
                                data[1] = read1_buffer[23];
                                data[2] = read1_buffer[22];
                                data[3] = read1_buffer[21];
                                AzimuthTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    AzimuthTextBlock.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                }));
                                data[0] = read1_buffer[28];
                                data[1] = read1_buffer[27];
                                data[2] = read1_buffer[26];
                                data[3] = read1_buffer[25];
                                TargetTextBock.Dispatcher.Invoke(new Action(delegate
                                {
                                    TargetTextBock.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                }));
                                data[0] = read1_buffer[32];
                                data[1] = read1_buffer[31];
                                data[2] = read1_buffer[30];
                                data[3] = read1_buffer[29];
                                AngleTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    AngleTextBlock.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                }));
                                #endregion
                                #region Time data
                                string TimeBuffer = Convert.ToUInt16(read1_buffer[49] << 8 | read1_buffer[50]).ToString("D04");
                                TimeBuffer += "-" + Convert.ToUInt16(read1_buffer[51] << 8 | read1_buffer[52]).ToString("D02");
                                TimeBuffer += "-" + Convert.ToUInt16(read1_buffer[53] << 8 | read1_buffer[54]).ToString("D02");
                                TimeBuffer += " " + Convert.ToUInt16(read1_buffer[55] << 8 | read1_buffer[56]).ToString("D02");
                                TimeBuffer += ":" + Convert.ToUInt16(read1_buffer[57] << 8 | read1_buffer[58]).ToString("D02");
                                TimeBuffer += ":" + Convert.ToUInt16(read1_buffer[59] << 8 | read1_buffer[60]).ToString("D02");
                                TimeTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    TimeTextBlock.Text = TimeBuffer;
                                }));
                                #endregion
                                WindSpeedTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    WindSpeedTextBlock.Text = Convert.ToUInt16(read1_buffer[69] << 8 | read1_buffer[70]).ToString() + "m/s";
                                }));
                                TemperatureTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    TemperatureTextBlock.Text = Convert.ToUInt16(read1_buffer[81] << 8 | read1_buffer[82]).ToString() + "℃";
                                }));
                                MotorCurTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    MotorCurTextBlock.Text = (Convert.ToUInt16(read1_buffer[83] << 8 | read1_buffer[84]) / 10.0).ToString("F01") + "A";
                                }));
                                #region Power
                                PVBuckVolTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    PVBuckVolTextBlock.Text = (Convert.ToUInt16(read1_buffer[113] << 8 | read1_buffer[114]) / 10.0).ToString("F01") + "V";
                                }));
                                PVStringCurTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    PVStringCurTextBlock.Text = (read1_buffer[38 * 2 + 3] / 10.0).ToString("F01") + "|" + (read1_buffer[38 * 2 + 3] / 10.0).ToString("F01") + "A"; ;
                                }));
                                BatCurTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    data[0] = read1_buffer[59 * 2 + 3 + 1];
                                    data[1] = read1_buffer[59 * 2 + 3];
                                    BatCurTextBlock.Text = (BitConverter.ToInt16(data, 0) / 100.0).ToString("F02") + "A";
                                }));
                                BatVolSocTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    BatVolSocTextBlock.Text = (Convert.ToUInt16(read1_buffer[58 * 2 + 3] << 8 | read1_buffer[58 * 2 + 3 + 1]) / 10.0).ToString("F01") + "V/" + read1_buffer[57 * 2 + 3].ToString() + "%";
                                }));
                                BatTempTextBlock.Dispatcher.Invoke(new Action(delegate
                                {
                                    BatTempTextBlock.Text = ((sbyte)read1_buffer[57 * 2 + 4]).ToString() + "℃";
                                }));
                                #endregion
                                #region charger
                                ushort PowerState = Convert.ToUInt16(read1_buffer[56 * 2 + 3] << 8 | read1_buffer[56 * 2 + 3 + 1]);
                                HeatIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & 1) > 0) HeatIndicator.Fill = GreenBrush;
                                    else HeatIndicator.Fill = DimGrayBrush;
                                }));
                                PreChargeIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 1)) > 0) PreChargeIndicator.Fill = GreenBrush;
                                    else PreChargeIndicator.Fill = DimGrayBrush;
                                }));
                                CCCVIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 2)) > 0) CCCVIndicator.Fill = GreenBrush;
                                    else CCCVIndicator.Fill = DimGrayBrush;
                                }));
                                TerminationIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 3)) > 0) TerminationIndicator.Fill = GreenBrush;
                                    else TerminationIndicator.Fill = DimGrayBrush;
                                }));
                                MaxTimeIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 4)) > 0) MaxTimeIndicator.Fill = RedBrush;
                                    else MaxTimeIndicator.Fill = DimGrayBrush;
                                }));
                                BatMissingIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 5)) > 0) BatMissingIndicator.Fill = RedBrush;
                                    else BatMissingIndicator.Fill = DimGrayBrush;
                                }));
                                BatShortIndicator.Dispatcher.Invoke(new Action(delegate
                                {
                                    if ((PowerState & (1 << 6)) > 0) BatShortIndicator.Fill = RedBrush;
                                    else BatShortIndicator.Fill = DimGrayBrush;
                                }));
                                #endregion
                                if (buttonstate == 1)
                                {
                                    buttonstate = 0;
                                    #region Astronomy parameter
                                    data[0] = read1_buffer[46 * 2 + 3 + 1];
                                    data[1] = read1_buffer[46 * 2 + 3];
                                    data[2] = read1_buffer[45 * 2 + 3 + 1];
                                    data[3] = read1_buffer[45 * 2 + 3];
                                    LongitudeTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LongitudeTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F04");
                                    }));
                                    data[0] = read1_buffer[48 * 2 + 3 + 1];
                                    data[1] = read1_buffer[48 * 2 + 3];
                                    data[2] = read1_buffer[47 * 2 + 3 + 1];
                                    data[3] = read1_buffer[47 * 2 + 3];
                                    LatitudeTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        LatitudeTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F04");
                                    }));
                                    data[0] = read1_buffer[50 * 2 + 3 + 1];
                                    data[1] = read1_buffer[50 * 2 + 3];
                                    data[2] = read1_buffer[49 * 2 + 3 + 1];
                                    data[3] = read1_buffer[49 * 2 + 3];
                                    TimeZoneTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        TimeZoneTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F04");
                                    }));
                                    #endregion
                                    #region Arry parameter
                                    data[0] = read1_buffer[90 * 2 + 3 + 1];
                                    data[1] = read1_buffer[90 * 2 + 3];
                                    data[2] = read1_buffer[89 * 2 + 3 + 1];
                                    data[3] = read1_buffer[89 * 2 + 3];
                                    ModuleWidthTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        ModuleWidthTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    data[0] = read1_buffer[92 * 2 + 3 + 1];
                                    data[1] = read1_buffer[92 * 2 + 3];
                                    data[2] = read1_buffer[91 * 2 + 3 + 1];
                                    data[3] = read1_buffer[91 * 2 + 3];
                                    PostSpaceTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        PostSpaceTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    data[0] = read1_buffer[94 * 2 + 3 + 1];
                                    data[1] = read1_buffer[94 * 2 + 3];
                                    data[2] = read1_buffer[93 * 2 + 3 + 1];
                                    data[3] = read1_buffer[93 * 2 + 3];
                                    TerrSlopeTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        TerrSlopeTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    #endregion
                                    #region Angle&Cur parameter
                                    data[0] = read1_buffer[52 * 2 + 3 + 1];
                                    data[1] = read1_buffer[52 * 2 + 3];
                                    data[2] = read1_buffer[51 * 2 + 3 + 1];
                                    data[3] = read1_buffer[51 * 2 + 3];
                                    AccuracyTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        AccuracyTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    data[0] = read1_buffer[44 * 2 + 3 + 1];
                                    data[1] = read1_buffer[44 * 2 + 3];
                                    data[2] = read1_buffer[43 * 2 + 3 + 1];
                                    data[3] = read1_buffer[43 * 2 + 3];
                                    MaxMotorCurTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        MaxMotorCurTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    #endregion
                                    #region ProAngle parameter
                                    data[0] = read1_buffer[104 * 2 + 3 + 1];
                                    data[1] = read1_buffer[104 * 2 + 3];
                                    data[2] = read1_buffer[103 * 2 + 3 + 1];
                                    data[3] = read1_buffer[103 * 2 + 3];
                                    ESoftLimitTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        ESoftLimitTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    data[0] = read1_buffer[106 * 2 + 3 + 1];
                                    data[1] = read1_buffer[106 * 2 + 3];
                                    data[2] = read1_buffer[105 * 2 + 3 + 1];
                                    data[3] = read1_buffer[105 * 2 + 3];
                                    WSoftLimitTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        WSoftLimitTextBox.Text = (BitConverter.ToSingle(data, 0)).ToString("F02");
                                    }));
                                    NightDockTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        NightDockTextBox.Text = Convert.ToUInt16(read1_buffer[111 * 2 + 3] << 8 | read1_buffer[111 * 2 + 3 + 1]).ToString();
                                    }));
                                    CleanAngleTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        CleanAngleTextBox.Text = Convert.ToUInt16(read1_buffer[84 * 2 + 3] << 8 | read1_buffer[84 * 2 + 3 + 1]).ToString();
                                    }));
                                    #endregion
                                    if (read1_buffer[85 * 2 + 3] == 0 && read1_buffer[85 * 2 + 3 + 1] == 0)
                                    {
                                        EastwardCheckBox.Dispatcher.Invoke(new Action(delegate
                                        {
                                            EastwardCheckBox.IsChecked = false;
                                        }));
                                    }
                                    else if (read1_buffer[85 * 2 + 3] == 0 && read1_buffer[85 * 2 + 3 + 1] == 1)
                                    {
                                        EastwardCheckBox.Dispatcher.Invoke(new Action(delegate
                                        {
                                            EastwardCheckBox.IsChecked = true;
                                        }));
                                    }

                                    #region ProWind parmaeter
                                    ProWindUpTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        ProWindUpTextBox.Text = Convert.ToUInt16(read1_buffer[99 * 2 + 3] << 8 | read1_buffer[99 * 2 + 3 + 1]).ToString();
                                    }));
                                    ProWindDownTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        ProWindDownTextBox.Text = Convert.ToUInt16(read1_buffer[101 * 2 + 3] << 8 | read1_buffer[101 * 2 + 3 + 1]).ToString();
                                    }));
                                    ProTimeUpTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        ProTimeUpTextBox.Text = Convert.ToUInt16(read1_buffer[100 * 2 + 3] << 8 | read1_buffer[100 * 2 + 3 + 1]).ToString();
                                    }));
                                    ProTimeDownTextBox.Dispatcher.Invoke(new Action(delegate
                                    {
                                        ProTimeDownTextBox.Text = Convert.ToUInt16(read1_buffer[102 * 2 + 3] << 8 | read1_buffer[102 * 2 + 3 + 1]).ToString();
                                    }));
                                    #endregion
                                    MessageBox.Show("读取成功，read successful !");
                                }
                            }
                            else if (bytes_to_read == 8)
                            {
                                if( read1_buffer[1] == 0x10 &&
                                    read1_buffer[2] == 0x00 &&
                                    read1_buffer[3] == 0x2b  )
                                {
                                    MessageBox.Show("修改成功，write successful !");
                                }                                
                            }
                        }
                    }

                }

            }
        }
        private void th_search()
        {
            while(true)
            {

            }
        }
        private void th_upgrade()
        {
            uint i = 0, j = 0, k = 0;
            while (true)
            {
                if (buttonstate == 1)
                {
                    FileInfo fi = new FileInfo(upgradefile.FileName);
                    long filelen = fi.Length;
                    FileStream fs = new FileStream(upgradefile.FileName, FileMode.Open);
                    byte[] buffer = new byte[255];

                    upgradeprogressbar.Dispatcher.Invoke(new Action(delegate
                    {
                        if (filelen - ((filelen / 256) * 256) > 0)
                        {
                            upgradeprogressbar.Maximum = (1 + (filelen / 256) + 1 + 1) * Convert.ToInt16(upgradecount.Text);
                        }
                        else
                        {
                            upgradeprogressbar.Maximum = (1 + (filelen / 256) + 1) * Convert.ToInt16(upgradecount.Text);
                        }
                        upgradeprogressbar.Value = 0;
                    }));

                    send1_buffer[0] = 0x00; send1_buffer[1] = 0x64; send1_buffer[2] = 0x00;
                    upgradedevicetype.Dispatcher.Invoke(new Action(delegate
                    {
                        send1_buffer[3] = Convert.ToByte(upgradedevicetype.Text);
                    }));
                    GetCRC16CheckCode(4, out send1_buffer[4], out send1_buffer[5], send1_buffer);
                    upgradecount.Dispatcher.Invoke(new Action(delegate
                    {
                        for (i = 0; i < Convert.ToInt16(upgradecount.Text); i++)
                        {
                            SerialComPort1.Write(send1_buffer, 0, 6);
                            upgradedelaytime.Dispatcher.Invoke(new Action(delegate
                            {
                                // Thread.Sleep(Convert.ToInt16(upgradedelaytime.Text));
                                Delay(Convert.ToInt16(upgradedelaytime.Text));
                            }));
                            upgradeprogressbar.Dispatcher.Invoke(new Action(delegate
                            {
                                upgradeprogressbar.Value++;
                            }));
                        }
                    }));

                    send1_buffer[0] = 0x00; send1_buffer[1] = 0x65;
                    byte[] sum = new byte[1];
                    sum[0] = 0xff;

                    for (j = 0; j < (filelen / 256); j++)
                    {
                        fs.Seek(j * 256, SeekOrigin.Begin);
                        fs.Read(send1_buffer, 4, 256);
                        if (j == 0)
                        {
                            for (k = 0; k < 16; k++)
                            {
                                send1_buffer[4 + k] = Convert.ToByte(sum[0] - send1_buffer[4 + k]);
                            }
                        }
                        send1_buffer[2] = Convert.ToByte(j >> 8); send1_buffer[3] = Convert.ToByte(j & 0xff);
                        GetCRC16CheckCode(260, out send1_buffer[260], out send1_buffer[261], send1_buffer);
                        upgradecount.Dispatcher.Invoke(new Action(delegate
                        {
                            for (i = 0; i < Convert.ToInt16(upgradecount.Text); i++)
                            {
                                SerialComPort1.Write(send1_buffer, 0, 262);
                                upgradedelaytime.Dispatcher.Invoke(new Action(delegate
                                {
                                   // Thread.Sleep(Convert.ToInt16(upgradedelaytime.Text));
                                    Delay(Convert.ToInt16(upgradedelaytime.Text));
                                }));
                                upgradeprogressbar.Dispatcher.Invoke(new Action(delegate
                                {
                                    upgradeprogressbar.Value++;
                                }));
                            }
                        }));
                    }
                    if (filelen - ((filelen / 256) * 256) > 0)
                    {
                        fs.Seek(j * 256, SeekOrigin.Begin);
                        fs.Read(send1_buffer, 4, Convert.ToInt16(filelen - ((filelen / 256) * 256)));
                        send1_buffer[2] = Convert.ToByte(j >> 8); send1_buffer[3] = Convert.ToByte(j & 0xff);
                        GetCRC16CheckCode(Convert.ToInt16(filelen - ((filelen / 256) * 256)) + 4, out send1_buffer[Convert.ToInt16(filelen - ((filelen / 256) * 256)) + 4], out send1_buffer[Convert.ToInt16(filelen - ((filelen / 256) * 256)) + 4 + 1], send1_buffer);
                        upgradecount.Dispatcher.Invoke(new Action(delegate
                        {
                            for (i = 0; i < Convert.ToInt16(upgradecount.Text); i++)
                            {
                                SerialComPort1.Write(send1_buffer, 0, Convert.ToInt16(filelen - ((filelen / 256) * 256)) + 4 + 2);
                                upgradedelaytime.Dispatcher.Invoke(new Action(delegate
                                {
                                   // Thread.Sleep(Convert.ToInt16(upgradedelaytime.Text));
                                    Delay(Convert.ToInt16(upgradedelaytime.Text));
                                }));
                                upgradeprogressbar.Dispatcher.Invoke(new Action(delegate
                                {
                                    upgradeprogressbar.Value++;
                                }));
                            }
                        }));

                        send1_buffer[0] = 0x00; send1_buffer[1] = 0x66;
                        send1_buffer[2] = Convert.ToByte(Convert.ToUInt16(j + 1) >> 8); send1_buffer[3] = Convert.ToByte(Convert.ToUInt16(j + 1) & 0xff);
                        GetCRC16CheckCode(4, out send1_buffer[4], out send1_buffer[5], send1_buffer);
                        upgradecount.Dispatcher.Invoke(new Action(delegate
                        {
                            for (i = 0; i < Convert.ToInt16(upgradecount.Text); i++)
                            {
                                SerialComPort1.Write(send1_buffer, 0, 6);
                                upgradedelaytime.Dispatcher.Invoke(new Action(delegate
                                {
                                    //Thread.Sleep(Convert.ToInt16(upgradedelaytime.Text));
                                    Delay(Convert.ToInt16(upgradedelaytime.Text));
                                }));
                                upgradeprogressbar.Dispatcher.Invoke(new Action(delegate
                                {
                                    upgradeprogressbar.Value++;
                                }));
                            }
                        }));
                    }
                    else
                    {
                        send1_buffer[0] = 0x00; send1_buffer[1] = 0x66;
                        send1_buffer[2] = Convert.ToByte(Convert.ToUInt16(j) >> 8); send1_buffer[3] = Convert.ToByte(Convert.ToUInt16(j) & 0xff);
                        GetCRC16CheckCode(4, out send1_buffer[4], out send1_buffer[5], send1_buffer);
                        upgradecount.Dispatcher.Invoke(new Action(delegate
                        {
                            for (i = 0; i < Convert.ToInt16(upgradecount.Text); i++)
                            {
                                SerialComPort1.Write(send1_buffer, 0, 6);
                                upgradedelaytime.Dispatcher.Invoke(new Action(delegate
                                {
                                    //Thread.Sleep(Convert.ToInt16(upgradedelaytime.Text));
                                    Delay(Convert.ToInt16(upgradedelaytime.Text));
                                }));
                                upgradeprogressbar.Dispatcher.Invoke(new Action(delegate
                                {
                                    upgradeprogressbar.Value++;
                                }));
                            }
                        }));
                    }
                    fs.Close();
                    buttonstate = 0;
                }
            }
        }
        private void th_scan()
        {
            while (true)
            {

            }
        }

        private void TabButon_Click(object sender, RoutedEventArgs e)
        {
            Button _button = sender as Button;

            if ((t_control.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_control.Abort();
                t_control.Join();
            }
            if ((t_scan.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_scan.Abort();
                t_scan.Join();
            }
            if ((t_upgrade.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_upgrade.Abort();
                t_upgrade.Join();
            }
            if ((t_search.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_search.Abort();
                t_search.Join();
            }
            if ((t_inicfg.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_inicfg.Abort();
                t_inicfg.Join();
            }
            if ((t_power.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_power.Abort();
                t_power.Join();
            }
            if ((t_onekey.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_onekey.Abort();
                t_onekey.Join();
            }
            switch (_button.Name)
            {
                #region INICFG Buton
                case "INICFGButton":
                    MainTabControl.SelectedIndex = 0;
                    INICFGButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_inicfg = new Thread(t_inicfgstart);
                        t_inicfg.Start();
                    }
                    POWERButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    ONESHOTButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion

                #region POWER Buton
                case "POWERButton":
                    MainTabControl.SelectedIndex = 1;
                    POWERButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_power = new Thread(t_powerstart);
                        t_power.Start();
                    }
                    INICFGButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    ONESHOTButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion

                #region CONTROL Buton
                case "CONTROLButton":
                    MainTabControl.SelectedIndex = 2;
                    CONTROLButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_control = new Thread(t_controlstart);
                        t_control.Start();
                    }
                    INICFGButton.Background = TabButtonLightGrayBrush;
                    POWERButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    ONESHOTButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion

                #region LORASEARCH Buton
                case "LORASEARCHButton":
                    MainTabControl.SelectedIndex = 7;
                    LORASEARCHButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_search = new Thread(t_searchstart);
                        t_search.Start();
                    }
                    INICFGButton.Background = TabButtonLightGrayBrush;
                    POWERButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    ONESHOTButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion

                #region UPGRADE Buton
                case "UPGRADEButton":
                    MainTabControl.SelectedIndex = 4;
                    UPGRADEButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_upgrade = new Thread(t_upgradestart);
                        t_upgrade.Start();
                    }
                    INICFGButton.Background = TabButtonLightGrayBrush;
                    POWERButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    ONESHOTButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion

                #region SCAN Buton
                case "SCANButton":
                    MainTabControl.SelectedIndex = 7;
                    SCANButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_scan = new Thread(t_scanstart);
                        t_scan.Start();
                   }
                    INICFGButton.Background = TabButtonLightGrayBrush;
                    POWERButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    ONESHOTButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion

                #region OneShotButon
                case "ONESHOTButton":
                    MainTabControl.SelectedIndex = 6;
                    ONESHOTButton.Background = TabButtonGrayBrush;
                    if (SerialComPort1.IsOpen == true)
                    {
                        t_onekey = new Thread(t_onekeystart);
                        t_onekey.Start();
                    }
                    INICFGButton.Background = TabButtonLightGrayBrush;
                    POWERButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    AboutButton.Background = TabButtonLightGrayBrush;
                    break;
                #endregion
                #region AboutButon
                case "AboutButton":
                    MainTabControl.SelectedIndex = 7;
                    AboutButton.Background = TabButtonGrayBrush;

                    INICFGButton.Background = TabButtonLightGrayBrush;
                    POWERButton.Background = TabButtonLightGrayBrush;
                    CONTROLButton.Background = TabButtonLightGrayBrush;
                    LORASEARCHButton.Background = TabButtonLightGrayBrush;
                    UPGRADEButton.Background = TabButtonLightGrayBrush;
                    SCANButton.Background = TabButtonLightGrayBrush;
                    break;
                    #endregion

            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Button _button = sender as Button;

            if (SerialComPort1.IsOpen == false) return;

            if( _button.Name != "ReadAllButton" &&
                _button.Name != "GoButton"      &&
                _button.Name != "WriteAllButton")
            {
                _button.Background = ButtonBlueBrush;
                _button.Foreground = TextWhiteBrush;
            }

            switch (_button.Name)
            {
                #region POWER LowPower Button 6
                case "P_LowPowerButton":
                    buttonstate = 6;
                    break;
                #endregion
                #region POWER RESET Button 14
                case "P_ResetButton":
                    buttonstate = 14;
                    break;
                #endregion
                #region POWER SLEEP Button 15
                case "P_SleepButton":
                    buttonstate = 15;
                    break;
                #endregion
                #region POWER changeid Button 17
                case "P_ChangeIDButton":
                    buttonstate = 17;
                    break;
                #endregion

                #region control read and write Button 1-2
                case "ReadAllButton":
                    buttonstate = 1;
                    break;
                case "WriteAllButton":
                    buttonstate = 2;
                    break;
                #endregion
                #region control mode Button 3-15
                case "AutoButton":
                    buttonstate = 3;
                    break;
                case "AIButton":
                    buttonstate = 4;
                    break;
                case "WindButton":
                    buttonstate = 5;
                    break;
                case "LowPowerButton":
                    buttonstate = 6;
                    break;
                case "SnowButton":
                    buttonstate = 7;
                    break;
                case "RainButton":
                    buttonstate = 8;
                    break;
                case "EastButton":
                    buttonstate = 9;
                    break;
                case "WestButton":
                    buttonstate = 10;
                    break;
                case "CleanButton":
                    buttonstate = 11;
                    break;
                case "CaliButton":
                    buttonstate = 12;
                    break;
                case "HorizontalButton":
                    buttonstate = 13;
                    break;
                case "C_ResetButton":
                    buttonstate = 14;
                    break;
                case "C_SleepButton":
                    buttonstate = 15;
                    break;
                #endregion
                #region control go Button 16
                case "GoButton":
                    if (AIAngleTextBox.Text == null)
                    {
                        MessageBox.Show("Angle error !");
                        buttonstate = 0;
                        break;
                    }
                    buttonstate = 16;
                    break;
                #endregion
                #region control changeid Button 17
                case "ChangeIDButton":
                    buttonstate = 17;
                    break;
                #endregion
                #region updatetime Button 18
                case "UpdateTimeButton":
                    buttonstate = 18;
                    break;
                #endregion
                #region upgradefileselect Button 1
                case "upgradefileselect":
                    upgradefile.Filter = "bin文件|*.bin";
                    upgradefile.Multiselect = false;
                    if (upgradefile.ShowDialog() == true)
                    {
                        MessageBox.Show(upgradefile.FileName);
                    }
                    else
                    {
                        MessageBox.Show("打开失败");
                        return;
                    }
                    upgradepath.Text = upgradefile.FileName;
                    buttonstate = 1;
                    break;
                #endregion
                default: break;
            }
        }

        private void Button_OneKey(object sender, RoutedEventArgs e)
        {
            Button _button = sender as Button;

            switch (_button.Name)
            {
                case "oneshotonekeyall":
                    buttonstate = 20;
                    break;
                case "oneshotonekeyeliminatedefect":
                    buttonstate = 0;// 21;
                    break;
                case "P_oneshotonekeyall":
                    buttonstate = 0;// 22;
                    break;
                case "P_oneshotonekeyeliminatedefect":
                    buttonstate = 0;// 23;
                    break;
                default: break;
            }
        }

        private void onekeyallPLcheckbox_Click(object sender, RoutedEventArgs e)                 //onekeyallPLcheckbox change onekeyIDtextbox Visibility      
        {
            if (onekeyallPLcheckbox.IsChecked == true)
            {
                onekeyIDtextbox.Visibility = Visibility.Hidden;
            }
            else
            {
                onekeyIDtextbox.Visibility = Visibility.Visible;
            }
        }

        private void LoraSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)     //lora select change,IsEnabled change
        {
            ComboBoxItem item = this.LoraSelect.SelectedItem as ComboBoxItem;

            if (item.Content.ToString() == "捷迅易联")
            {
                ybt_AirRate.IsEnabled = false;
                ybt_Baudrate.IsEnabled = false;
                ybt_ChannelRSSI.IsEnabled = false;
                ybt_Character.IsEnabled = false;
                ybt_Cycle.IsEnabled = false;
                ybt_DataRSSI.IsEnabled = false;
                ybt_LBTEnable.IsEnabled = false;
                ybt_PacketSize.IsEnabled = false;
                ybt_Parity.IsEnabled = false;
                ybt_Power.IsEnabled = false;
                ybt_RelayEnable.IsEnabled = false;
                ybt_TransforMode.IsEnabled = false;
                ybt_DeviceId.IsEnabled = false;
                ybt_FreqChannel.IsEnabled = false;
                ybt_NetId.IsEnabled = false;
                ybt_key.IsEnabled = false;
                ybt_module.IsEnabled = false;

                jxyl_Baudrate.IsEnabled = true;
                jxyl_BW.IsEnabled = true;
                jxyl_factor.IsEnabled = true;
                jxyl_Freq.IsEnabled = true;
                jxyl_Mode.IsEnabled = true;
                jxyl_NetId.IsEnabled = true;
                jxyl_NodeId1.IsEnabled = true;
                jxyl_NodeId2.IsEnabled = true;
                jxyl_Parity.IsEnabled = true;
                jxyl_Power.IsEnabled = true;
                jxyl_module.IsEnabled = true;
            }
            else if (item.Content.ToString() == "易佰特")
            {
                jxyl_Baudrate.IsEnabled = false;
                jxyl_BW.IsEnabled = false;
                jxyl_factor.IsEnabled = false;
                jxyl_Freq.IsEnabled = false;
                jxyl_Mode.IsEnabled = false;
                jxyl_NetId.IsEnabled = false;
                jxyl_NodeId1.IsEnabled = false;
                jxyl_NodeId2.IsEnabled = false;
                jxyl_Parity.IsEnabled = false;
                jxyl_Power.IsEnabled = false;
                jxyl_module.IsEnabled = false;

                ybt_AirRate.IsEnabled = true;
                ybt_Baudrate.IsEnabled = true;
                ybt_ChannelRSSI.IsEnabled = true;
                ybt_Character.IsEnabled = true;
                ybt_Cycle.IsEnabled = true;
                ybt_DataRSSI.IsEnabled = true;
                ybt_LBTEnable.IsEnabled = true;
                ybt_PacketSize.IsEnabled = true;
                ybt_Parity.IsEnabled = true;
                ybt_Power.IsEnabled = true;
                ybt_RelayEnable.IsEnabled = true;
                ybt_TransforMode.IsEnabled = true;
                ybt_DeviceId.IsEnabled = true;
                ybt_FreqChannel.IsEnabled = true;
                ybt_NetId.IsEnabled = true;
                ybt_key.IsEnabled = true;
                ybt_module.IsEnabled = true;
            }
        }

        private void ButtonINI_Click(object sender, RoutedEventArgs e)                           //INI BUTTON click
        {
            Button _button = sender as Button;

            if (_button.Name == "INICreateButton")
            {
                #region INICreateButton
                if (LoraSelect.SelectedIndex == -1)
                {
                    MessageBox.Show("LORA未选择，LORA NO SELECT");
                    return;
                }
                if (MessageBox.Show("参数已检查，已检查点确定，再次检查点取消，Confirmed parameters，YES click OK, Confirm again click Cancel",
                    "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.Cancel)
                {
                    return;
                }

                //创建一个保存文件式的对话框
                SaveFileDialog sfd = new SaveFileDialog();
                //设置这个对话框的起始保存路径
                sfd.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;// @"D:\";
                                                                             //设置保存的文件的类型，注意过滤器的语法
                sfd.Filter = "ini文件|*.ini";
                //调用ShowDialog()方法显示该对话框，该方法的返回值代表用户是否点击了确定按钮
                if (sfd.ShowDialog() == true)
                {
                    MessageBox.Show("保存成功");
                }
                else
                {
                    MessageBox.Show("取消保存");
                    return;
                }
                //LORA选择

                ComboBoxItem item = this.LoraSelect.SelectedItem as ComboBoxItem;
                string DeviceSelect = item.Content.ToString(); ;
                if (DeviceSelect == "易佰特")
                {
                    INI.WriteIni("LORA-seclet", "lora", "易佰特", sfd.FileName);
                    //易佰特
                    INI.WriteIni("LORA-ybt", "baudrate", ybt_Baudrate.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "parity", ybt_Parity.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "airrate", ybt_AirRate.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "packetsize", ybt_PacketSize.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "character", ybt_Character.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "cycle", ybt_Cycle.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "power", ybt_Power.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "transformode", ybt_TransforMode.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "relayenable", ybt_RelayEnable.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "lbtenable", ybt_LBTEnable.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "datarssi", ybt_DataRSSI.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "channelrssi", ybt_ChannelRSSI.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-ybt", "deviceid", ybt_DeviceId.Text, sfd.FileName);
                    INI.WriteIni("LORA-ybt", "freqchannel", ybt_FreqChannel.Text, sfd.FileName);
                    INI.WriteIni("LORA-ybt", "netid", ybt_NetId.Text, sfd.FileName);
                    INI.WriteIni("LORA-ybt", "key", ybt_key.Text, sfd.FileName);
                    INI.WriteIni("LORA-ybt", "module", ybt_module.SelectedIndex.ToString(), sfd.FileName);
                }
                else if (DeviceSelect == "捷迅易联")
                {
                    INI.WriteIni("LORA-seclet", "lora", "捷迅易联", sfd.FileName);
                    //捷迅易联
                    INI.WriteIni("LORA-jxyl", "freq", jxyl_Freq.Text, sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "factor", jxyl_factor.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "mode", jxyl_Mode.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "bw", jxyl_BW.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "nodeid1", jxyl_NodeId1.Text, sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "nodeid2", jxyl_NodeId2.Text, sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "netid", jxyl_NetId.Text, sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "power", jxyl_Power.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "baudrate", jxyl_Baudrate.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "parity", jxyl_Parity.SelectedIndex.ToString(), sfd.FileName);
                    INI.WriteIni("LORA-jxyl", "module", jxyl_module.SelectedIndex.ToString(), sfd.FileName);
                }
                INI.WriteIni("Power", "id", P_DeviceID.Text, sfd.FileName);
                if (P_IDMode1CheckBox.IsChecked == true)
                {
                    INI.WriteIni("Power", "idmode", "-", sfd.FileName);
                }
                else if (P_IDMode1CheckBox.IsChecked == false)
                {
                    INI.WriteIni("Power", "idmode", "+", sfd.FileName);
                }
                INI.WriteIni("Power", "number", P_IDMode.Text, sfd.FileName);
                INI.WriteIni("Power", "idmax", P_IDMAX.Text, sfd.FileName);

                INI.WriteIni("Control", "id", C_DeviceID.Text, sfd.FileName);
                if (C_IDMode1CheckBox.IsChecked == false)
                {
                    INI.WriteIni("Control", "idmode", "+", sfd.FileName);
                }
                else if (C_IDMode1CheckBox.IsChecked == true)
                {
                    INI.WriteIni("Control", "idmode", "-", sfd.FileName);
                }
                INI.WriteIni("Control", "number", C_IDMode.Text, sfd.FileName);
                INI.WriteIni("Control", "idmax", C_IDMAX.Text, sfd.FileName);

                INI.WriteIni("Control", "longitude", INILongitudeTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "latitude", INILatitudeTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "timezone", INITimeZoneTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "modulewidth", INIModuleWidthTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "postspace", INIPostSpaceTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "terrslope", INITerrSlopeTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "accuracy", INIAccuracyTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "maxmotorcur", INIMaxMotorCurTextBox.Text, sfd.FileName);
                if (INIEastwardCheckBox.IsChecked == false)
                {
                    INI.WriteIni("Control", "orientation", "east", sfd.FileName);
                }
                else if (INIEastwardCheckBox.IsChecked == true)
                {
                    INI.WriteIni("Control", "orientation", "west", sfd.FileName);
                }
                INI.WriteIni("Control", "ESoftLimit", INIESoftLimitTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "WSoftLimit", INIWSoftLimitTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "CleanAngle", INICleanAngleTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "NightDock", ININightDockTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "ProWindUp", INIProWindUpTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "ProWindDown", INIProWindDownTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "ProTimeUp", INIProTimeUpTextBox.Text, sfd.FileName);
                INI.WriteIni("Control", "ProTimeDown", INIProTimeDownTextBox.Text, sfd.FileName);
                #endregion
                return;
            }
            else if (_button.Name == "INIOpenButton")
            {
                #region INIOpenButton
                OpenFileDialog sfd = new OpenFileDialog();
                sfd.Filter = "ini文件|*.ini";
                sfd.Multiselect = false;

                if (sfd.ShowDialog() == true)
                {
                    MessageBox.Show(sfd.FileName);
                }
                else
                {
                    MessageBox.Show("打开失败");
                    return;
                }
                if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "易佰特")
                {
                    LoraSelect.SelectedIndex = 0;
                    ybt_Baudrate.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "baudrate", sfd.FileName));
                    ybt_Parity.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "parity", sfd.FileName));
                    ybt_AirRate.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "airrate", sfd.FileName));
                    ybt_PacketSize.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "packetsize", sfd.FileName));
                    ybt_Character.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "character", sfd.FileName));
                    ybt_Cycle.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "cycle", sfd.FileName));
                    ybt_Power.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "power", sfd.FileName));
                    ybt_TransforMode.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "transformode", sfd.FileName));
                    ybt_RelayEnable.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "relayenable", sfd.FileName));
                    ybt_LBTEnable.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "lbtenable", sfd.FileName));
                    ybt_DataRSSI.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "datarssi", sfd.FileName));
                    ybt_ChannelRSSI.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "channelrssi", sfd.FileName));
                    ybt_DeviceId.Text = INI.ReadIni("LORA-ybt", "deviceid", sfd.FileName);
                    ybt_FreqChannel.Text = INI.ReadIni("LORA-ybt", "freqchannel", sfd.FileName);
                    ybt_NetId.Text = INI.ReadIni("LORA-ybt", "netid", sfd.FileName);
                    ybt_key.Text = INI.ReadIni("LORA-ybt", "key", sfd.FileName);
                    ybt_module.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-ybt", "module", sfd.FileName));

                    jxyl_Freq.IsEnabled = false;
                    jxyl_factor.IsEnabled = false;
                    jxyl_Mode.IsEnabled = false;
                    jxyl_BW.IsEnabled = false;
                    jxyl_NodeId1.IsEnabled = false;
                    jxyl_NodeId2.IsEnabled = false;
                    jxyl_NetId.IsEnabled = false;
                    jxyl_Power.IsEnabled = false;
                    jxyl_Baudrate.IsEnabled = false;
                    jxyl_Parity.IsEnabled = false;
                    jxyl_module.IsEnabled = false;
                }
                else if (INI.ReadIni("LORA-seclet", "lora", sfd.FileName) == "捷迅易联")
                {
                    LoraSelect.SelectedIndex = 1;
                    jxyl_Freq.Text = INI.ReadIni("LORA-jxyl", "freq", sfd.FileName);
                    jxyl_factor.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "factor", sfd.FileName));
                    jxyl_Mode.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "mode", sfd.FileName));
                    jxyl_BW.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "bw", sfd.FileName));
                    jxyl_NodeId1.Text = INI.ReadIni("LORA-jxyl", "nodeid1", sfd.FileName);
                    jxyl_NodeId2.Text = INI.ReadIni("LORA-jxyl", "nodeid2", sfd.FileName);
                    jxyl_NetId.Text = INI.ReadIni("LORA-jxyl", "netid", sfd.FileName);
                    jxyl_Power.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "power", sfd.FileName));
                    jxyl_Baudrate.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "baudrate", sfd.FileName));
                    jxyl_Parity.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "parity", sfd.FileName));
                    jxyl_module.SelectedIndex = Convert.ToInt16(INI.ReadIni("LORA-jxyl", "module", sfd.FileName));

                    ybt_Baudrate.IsEnabled = false;
                    ybt_Parity.IsEnabled = false;
                    ybt_AirRate.IsEnabled = false;
                    ybt_PacketSize.IsEnabled = false;
                    ybt_Character.IsEnabled = false;
                    ybt_Cycle.IsEnabled = false;
                    ybt_Power.IsEnabled = false;
                    ybt_TransforMode.IsEnabled = false;
                    ybt_RelayEnable.IsEnabled = false;
                    ybt_LBTEnable.IsEnabled = false;
                    ybt_DataRSSI.IsEnabled = false;
                    ybt_ChannelRSSI.IsEnabled = false;
                    ybt_DeviceId.IsEnabled = false;
                    ybt_FreqChannel.IsEnabled = false;
                    ybt_NetId.IsEnabled = false;
                    ybt_key.IsEnabled = false;
                    ybt_module.IsEnabled = false;
                }

                P_DeviceID.Text = INI.ReadIni("Power", "id", sfd.FileName);
                if (INI.ReadIni("Power", "idmode", sfd.FileName) == "+")
                {
                    P_IDMode1CheckBox.IsChecked = false;
                }
                else if (INI.ReadIni("Power", "idmode", sfd.FileName) == "-")
                {
                    P_IDMode1CheckBox.IsChecked = true;
                }
                P_IDMode.Text = INI.ReadIni("Power", "number", sfd.FileName);
                P_IDMAX.Text = INI.ReadIni("Power", "idmax", sfd.FileName);

                C_DeviceID.Text = INI.ReadIni("Control", "id", sfd.FileName);
                if (INI.ReadIni("Control", "idmode", sfd.FileName) == "+")
                {
                    C_IDMode1CheckBox.IsChecked = false;
                }
                else if (INI.ReadIni("Control", "idmode", sfd.FileName) == "-")
                {
                    C_IDMode1CheckBox.IsChecked = true;
                }
                C_IDMode.Text = INI.ReadIni("Control", "number", sfd.FileName);
                C_IDMAX.Text = INI.ReadIni("Control", "idmax", sfd.FileName);

                INILongitudeTextBox.Text = INI.ReadIni("Control", "longitude", sfd.FileName);
                INILatitudeTextBox.Text = INI.ReadIni("Control", "latitude", sfd.FileName);
                INITimeZoneTextBox.Text = INI.ReadIni("Control", "timezone", sfd.FileName);
                INIModuleWidthTextBox.Text = INI.ReadIni("Control", "modulewidth", sfd.FileName);
                INIPostSpaceTextBox.Text = INI.ReadIni("Control", "postspace", sfd.FileName);
                INITerrSlopeTextBox.Text = INI.ReadIni("Control", "terrslope", sfd.FileName);
                INIAccuracyTextBox.Text = INI.ReadIni("Control", "accuracy", sfd.FileName);
                INIMaxMotorCurTextBox.Text = INI.ReadIni("Control", "maxmotorcur", sfd.FileName);

                if (INI.ReadIni("Control", "orientation", sfd.FileName) == "east")
                {
                    INIEastwardCheckBox.IsChecked = false;
                }
                else if (INI.ReadIni("Control", "orientation", sfd.FileName) == "west")
                {
                    INIEastwardCheckBox.IsChecked = true;
                }

                INIESoftLimitTextBox.Text = INI.ReadIni("Control", "ESoftLimit", sfd.FileName);
                INIWSoftLimitTextBox.Text = INI.ReadIni("Control", "WSoftLimit", sfd.FileName);
                INICleanAngleTextBox.Text = INI.ReadIni("Control", "CleanAngle", sfd.FileName);
                ININightDockTextBox.Text = INI.ReadIni("Control", "NightDock", sfd.FileName);
                INIProWindUpTextBox.Text = INI.ReadIni("Control", "ProWindUp", sfd.FileName);
                INIProWindDownTextBox.Text = INI.ReadIni("Control", "ProWindDown", sfd.FileName);
                INIProTimeUpTextBox.Text = INI.ReadIni("Control", "ProTimeUp", sfd.FileName);
                INIProTimeDownTextBox.Text = INI.ReadIni("Control", "ProTimeDown", sfd.FileName);
                modifypath = sfd.FileName;
                #endregion
                return;
            }
            else if (_button.Name == "INIModifyButton")
            {
                #region INIModifyButton
                if (MessageBox.Show(modifypath + "  修改确认，modify confirm",
                    "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.Cancel)
                {
                    return;
                }
                ComboBoxItem item = this.LoraSelect.SelectedItem as ComboBoxItem;
                string DeviceSelect = item.Content.ToString();
                if (DeviceSelect == "易佰特")
                {
                    INI.WriteIni("LORA-seclet", "lora", "易佰特", modifypath);
                    //易佰特
                    INI.WriteIni("LORA-ybt", "baudrate", ybt_Baudrate.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "parity", ybt_Parity.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "airrate", ybt_AirRate.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "packetsize", ybt_PacketSize.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "character", ybt_Character.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "cycle", ybt_Cycle.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "power", ybt_Power.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "transformode", ybt_TransforMode.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "relayenable", ybt_RelayEnable.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "lbtenable", ybt_LBTEnable.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "datarssi", ybt_DataRSSI.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "channelrssi", ybt_ChannelRSSI.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-ybt", "deviceid", ybt_DeviceId.Text, modifypath);
                    INI.WriteIni("LORA-ybt", "freqchannel", ybt_FreqChannel.Text, modifypath);
                    INI.WriteIni("LORA-ybt", "netid", ybt_NetId.Text, modifypath);
                    INI.WriteIni("LORA-ybt", "key", ybt_key.Text, modifypath);
                    INI.WriteIni("LORA-ybt", "module", ybt_module.SelectedIndex.ToString(), modifypath);
                }
                else if (DeviceSelect == "捷迅易联")
                {
                    INI.WriteIni("LORA-seclet", "lora", "捷迅易联", modifypath);
                    //捷迅易联
                    INI.WriteIni("LORA-jxyl", "freq", jxyl_Freq.Text, modifypath);
                    INI.WriteIni("LORA-jxyl", "factor", jxyl_factor.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-jxyl", "mode", jxyl_Mode.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-jxyl", "bw", jxyl_BW.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-jxyl", "nodeid1", jxyl_NodeId1.Text, modifypath);
                    INI.WriteIni("LORA-jxyl", "nodeid2", jxyl_NodeId2.Text, modifypath);
                    INI.WriteIni("LORA-jxyl", "netid", jxyl_NetId.Text, modifypath);
                    INI.WriteIni("LORA-jxyl", "power", jxyl_Power.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-jxyl", "baudrate", jxyl_Baudrate.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-jxyl", "parity", jxyl_Parity.SelectedIndex.ToString(), modifypath);
                    INI.WriteIni("LORA-jxyl", "module", jxyl_module.SelectedIndex.ToString(), modifypath);
                }
                INI.WriteIni("Power", "id", P_DeviceID.Text, modifypath);
                if (P_IDMode1CheckBox.IsChecked == false)
                {
                    INI.WriteIni("Power", "idmode", "+", modifypath);
                }
                else if (P_IDMode1CheckBox.IsChecked == true)
                {
                    INI.WriteIni("Power", "idmode", "-", modifypath);
                }
                INI.WriteIni("Power", "number", P_IDMode.Text, modifypath);
                INI.WriteIni("Power", "idmax", P_IDMAX.Text, modifypath);

                INI.WriteIni("Control", "id", C_DeviceID.Text, modifypath);
                if (C_IDMode1CheckBox.IsChecked == false)
                {
                    INI.WriteIni("Control", "idmode", "+", modifypath);
                }
                else if (C_IDMode1CheckBox.IsChecked == true)
                {
                    INI.WriteIni("Control", "idmode", "-", modifypath);
                }
                INI.WriteIni("Control", "number", C_IDMode.Text, modifypath);
                INI.WriteIni("Control", "idmax", C_IDMAX.Text, modifypath);

                INI.WriteIni("Control", "longitude", INILongitudeTextBox.Text, modifypath);
                INI.WriteIni("Control", "latitude", INILatitudeTextBox.Text, modifypath);
                INI.WriteIni("Control", "timezone", INITimeZoneTextBox.Text, modifypath);
                INI.WriteIni("Control", "modulewidth", INIModuleWidthTextBox.Text, modifypath);
                INI.WriteIni("Control", "postspace", INIPostSpaceTextBox.Text, modifypath);
                INI.WriteIni("Control", "terrslope", INITerrSlopeTextBox.Text, modifypath);
                INI.WriteIni("Control", "accuracy", INIAccuracyTextBox.Text, modifypath);
                INI.WriteIni("Control", "maxmotorcur", INIMaxMotorCurTextBox.Text, modifypath);
                if (INIEastwardCheckBox.IsChecked == false)
                {
                    INI.WriteIni("Control", "orientation", "east", modifypath);
                }
                else if (INIEastwardCheckBox.IsChecked == true)
                {
                    INI.WriteIni("Control", "orientation", "west", modifypath);
                }
                INI.WriteIni("Control", "ESoftLimit", INIESoftLimitTextBox.Text, modifypath);
                INI.WriteIni("Control", "WSoftLimit", INIWSoftLimitTextBox.Text, modifypath);
                INI.WriteIni("Control", "CleanAngle", INICleanAngleTextBox.Text, modifypath);
                INI.WriteIni("Control", "NightDock", ININightDockTextBox.Text, modifypath);
                INI.WriteIni("Control", "ProWindUp", INIProWindUpTextBox.Text, modifypath);
                INI.WriteIni("Control", "ProWindDown", INIProWindDownTextBox.Text, modifypath);
                INI.WriteIni("Control", "ProTimeUp", INIProTimeUpTextBox.Text, modifypath);
                INI.WriteIni("Control", "ProTimeDown", INIProTimeDownTextBox.Text, modifypath);
                MessageBox.Show("修改成功，Modified successfully");
                #endregion
                return;
            }
        }


        private void WindowX_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            buttonstate = 0;
            if ((t_control.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_control.Abort();
                t_control.Join();
            }
            if ((t_scan.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_scan.Abort();
                t_scan.Join();
            }
            if ((t_upgrade.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_upgrade.Abort();
                t_upgrade.Join();
            }
            if ((t_search.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_search.Abort();
                t_search.Join();
            }
            if ((t_inicfg.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_inicfg.Abort();
                t_inicfg.Join();
            }
            if ((t_power.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_power.Abort();
                t_power.Join();
            }
            if ((t_onekey.ThreadState & ThreadState.Unstarted) != 0) { }
            else
            {
                t_onekey.Abort();
                t_onekey.Join();
            }
            System.Environment.Exit(0);
        }

        private void WindowX_Loaded(object sender, RoutedEventArgs e)
        {
            if (INI.ReadIni("CONF", "language", AppDomain.CurrentDomain.BaseDirectory + "ONEFORALL.ini") == "English")
            {
                LanguageSelect.SelectedIndex = 0;
            }
            else if (INI.ReadIni("CONF", "language", AppDomain.CurrentDomain.BaseDirectory + "ONEFORALL.ini") == "简体中文")
            {
                LanguageSelect.SelectedIndex = 1;
            }

            #region tabcontrol初始化序号
            MainTabControl.SelectedIndex = 0;
            #endregion
            #region 控件初始化
            //ParameterTextBoxInit();
            //AlarmIndicaterInit();
            //BatteryChargerInit();
            #endregion
            ScanPortList(true);
            threadinit();

        }

        private void LanguageSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem item = this.LanguageSelect.SelectedItem as ComboBoxItem;
            INI.WriteIni("CONF", "language", item.Content.ToString(), AppDomain.CurrentDomain.BaseDirectory + "ONEFORALL.ini");
        }

        public static class DispatcherHelper
        {
            [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
            public static void DoEvents()
            {
                DispatcherFrame frame = new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrames), frame);
                try { Dispatcher.PushFrame(frame); }
                catch (InvalidOperationException) { }
            }
            private static object ExitFrames(object frame)
            {
                ((DispatcherFrame)frame).Continue = false;
                return null;
            }
        }

        public static void Delay(int mm)
        {
            DateTime current = DateTime.Now;
            while (current.AddMilliseconds(mm) > DateTime.Now)
            {
                DispatcherHelper.DoEvents();
            }
            return;
        }

    }
}
