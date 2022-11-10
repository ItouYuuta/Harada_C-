using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//追加
using System.IO.Ports;
using System.Management;//デバイス名も所得したいため
using System.Diagnostics;
using System.Configuration;

namespace 原田車両設計電池パック制御ユニット_PCアプリケーション
{
    public partial class Form1 : Form
    {
        //int judenbuttonFlg = 0;                                   //充電ボタンフラグ　1:押されている　0:押されていない
        //int kyuudenbuttonFlg = 0;                               //給電ボタンフラグ　1:押されている　0:押されていない
        bool judenbuttonFlg;                                          //充電ボタンフラグ　true:押されている　false:押されていない 既定値は false 
        bool kyuudenbuttonFlg;                                     //給電ボタンフラグ　true:押されている　false:押されていない 既定値は false 
        bool ijoumogibuttonFlg = false;                         //異常模擬ボタンフラグ　true:押されている　false:押されていない 既定値は false 
        //定義＋インスタンス生成
        SerialPort SrlPrt = new SerialPort();

        private delegate void Delegate_RcvDataToTextBox(string data);
        //送信バイトの単位作成
        byte[] sendByte = new byte[8];
       
        public Form1()
        {
            //初期化時にPCに接続されているCOMポート名をリストする
            InitializeComponent();

            //使用可能ポート取得
            string[] ports = SerialPort.GetPortNames(); //使用可能ポート取得
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

            foreach (string port in ports)
            {
                label2.Text = (port);
            }
        }
        //アプリ起動時の初期設定
        private void Form1_Load(object sender, EventArgs e)
        {

#if DEBUG//デバッグ時だけ実行
            this.Size = new Size(1280, 720); //フォーム１のサイズ
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            sndTextBox.Text = "";
            textBox2.Visible = true;//true:表示 false:非表示
            textBox3.Visible = true;//true:表示 false:非表示
            textBox4.Visible = true;//true:表示 false:非表示
            受信1バイト目.Visible = true;//true:表示 false:非表示
            受信2バイト目.Visible = true;//true:表示 false:非表示
            受信データ.Visible = true;//true:表示 false:非表示
            送信データ.Visible = true;//true:表示 false:非表示
            ただいまのポート.Visible = true;//true:表示 false:非表示
            ログ.Visible = true;//true:表示 false:非表示
#else//デバッグ以外で実行
            this.Size = new Size(720, 480);//フォーム１のサイズ
            textBox2.Visible = false;//true:表示 false:非表示
            textBox3.Visible = false;//true:表示 false:非表示
            textBox4.Visible = false;//true:表示 false:非表示
            受信1バイト目.Visible = false;//true:表示 false:非表示
            受信2バイト目.Visible = false;//true:表示 false:非表示
            受信データ.Visible = false;//true:表示 false:非表示
            送信データ.Visible = false;//true:表示 false:非表示
            ただいまのポート.Visible = false;//true:表示 false:非表示
            ログ.Visible = false;//true:表示 false:非表示
#endif//常に実行
            //各ボタン、テキストの文字設定
            lblGenzaiJyokyo.Text = "OFF";
            rcvTextBox.Text = "";
            //送信バイト初期値
            sendByte[0] = 0x50;
            sendByte[1] = 0x0a;
            sendByte[2] = 0x00;
            sendByte[3] = 0x00;
            sendByte[4] = 0x00;
            sendByte[5] = 0x00;
            sendByte[6] = 0x0d;
            sendByte[7] = 0x0a;
            if (serialPort1.IsOpen == true)
            {
                // シリアルポートをクローズする.
                serialPort1.Close();
                //ポート開放処理
                potokaihou();
            }
            else
            {
                //ポート開放処理
                potokaihou();
            }
        }

        //ポート開放処理
        private void potokaihou()
        {
            // シリアルポートをオープンする.
            serialPort1.BaudRate = int.Parse(ConfigurationManager.AppSettings["BaudRate"]);    //115200
            serialPort1.Parity = (Parity)int.Parse(ConfigurationManager.AppSettings["Parity"]);        //Parity.None;
            serialPort1.DataBits = int.Parse(ConfigurationManager.AppSettings["DataBits"]);    //8;
            serialPort1.StopBits = (StopBits)int.Parse(ConfigurationManager.AppSettings["StopBits"]);    //StopBits.One;
            serialPort1.Handshake = (Handshake)int.Parse(ConfigurationManager.AppSettings["Handshake"]);   //Handshake.None;
            serialPort1.PortName = ConfigurationManager.AppSettings["PortName"];
            try
            {

                // シリアルポートをオープンする.
                serialPort1.Open();

                //受信バッファをクリアする
                serialPort1.DiscardInBuffer();

                // ボタンの表示を[接続]から[切断]に変える.
                setuzokubutton.Text = "切断";
                // メッセージボックスを表示
                //MessageBox.Show("ポート接続完了", "お知らせ");

                int index = comboBox1.Items.IndexOf(ConfigurationManager.AppSettings["PortName"]); // 「COM3」のインデックスを取得する。
                comboBox1.SelectedIndex = index; // 「COM3」を指定状態にする。

                //コンボボックスの外部からの入力 true:有効 false:無効化
                comboBox1.Enabled = false;
            }
            catch (Exception)
            {
                // メッセージボックスを表示(アイコン付き)
                MessageBox.Show("ポート接続失敗\r\nポート接続は手動で接続して下さい。",
                    "お知らせ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                Porttoji();
            }
            Timersettei();
        }

        //接続ボタンを押した際、接続されていなかったら接続する
        private void setuzokubutton_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen == true)
            {
                //シリアルポートを閉じる
                serialPort1.Close();

                //ボタンの表示を[切断]から[接続]に変える
                setuzokubutton.Text = "接続";

                //コンボボックスの外部からの入力 true:有効 false:無効化
                comboBox1.Enabled = true;
            }
            else
            {
                //シリアルポートを開く
                serialPort1.BaudRate = 115200;//115200
                serialPort1.Parity = Parity.None;
                serialPort1.DataBits = 8;
                serialPort1.StopBits = StopBits.One;
                serialPort1.Handshake = Handshake.None;
                //使うポート名をcomboBoxで選択されたものにする
                serialPort1.PortName = comboBox1.SelectedItem.ToString();
                try
                {

                    //シリアルポートを開く
                    serialPort1.Open();

                    //受信バッファをクリアする
                    serialPort1.DiscardInBuffer();

                    //ボタンの表示を[接続]から[切断]に変える
                    setuzokubutton.Text = "切断";

                    //テキストボックスのログをクリア
                    rcvTextBox.Clear();

                    //コンボボックスの外部からの入力 true:有効 false:無効化
                    comboBox1.Enabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Porttoji();
                }
            }
#if DEBUG//デバッグ時だけ実行
            //textBox1に接続/切断したログを残す
            if (serialPort1.IsOpen == true)
            {
                //textBox1.Text = "接続しました。";
                textBox1.AppendText("接続しました。\r\n");
            }
            else
            {
                //textBox1.Text = "切断しました。";
                textBox1.AppendText("切断しました。\r\n");
            }
#endif//常に実行
            Setuzokukousin();
            Timersettei();
        }

        //ポート設定のドロップダウンを開く度に使用可能なポートを更新
        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            //string[] ports = SerialPort.GetPortNames();
            //ポート情報のクリア
            comboBox1.Items.Clear();
            //使用可能ポート取得
            string[] ports = SerialPort.GetPortNames();
            string com = "COM3";
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            comboBox1.SelectedIndex = 1;
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;
            foreach (string port in ports)
            {
                label2.Text = (port);
            }
        }

        //ウィンドウを閉じる前にポートを閉じる
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Porttoji();
        }

        //受信イベント02
        delegate void SetTextCallback(string text);
        private void Response(string responseData)
        {
            byte[] responseDataBytes = System.Text.Encoding.ASCII.GetBytes(responseData);

            if (rcvTextBox.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(Response);
                BeginInvoke(d, new object[] { responseData });
            }
            else
            {
                rcvTextBox.AppendText(responseData + "\r\n");

                try
                {

                    //---------------------
                    //1byte目を２進数で抽出
                    //---------------------
                    string responseDataByte_One = Convert.ToString(responseDataBytes[1], 2).PadLeft(8, '0');
                    int len2 = responseDataByte_One.Length;
                    txtResponseDataByte_One.Text = responseDataByte_One;

                    try
                    {
                        string responseDataByte_One_Bit0 = responseDataByte_One.Substring(len2 - 1, 1); //1byte目0bit目
                        string responseDataByte_One_Bit1 = responseDataByte_One.Substring(len2 - 2, 1); //1byte目1bit目
                        string responseDataByte_One_Bit2 = responseDataByte_One.Substring(len2 - 3, 1); //1byte目2bit目(終了フラグ)
                        string responseDataByte_One_Bit3 = responseDataByte_One.Substring(len2 - 4, 1); //1byte目3bit目(終了フラグ)
                        string responseDataByte_One_Bit4 = responseDataByte_One.Substring(len2 - 5, 1);//1byte目4bit目
                        string responseDataByte_One_Bit5 = responseDataByte_One.Substring(len2 - 6, 1); //1byte目5bit目
                        string responseDataByte_One_Bit6 = responseDataByte_One.Substring(len2 - 7, 1);//1byte目6bit目
                        string responseDataByte_One_Bit7 = responseDataByte_One.Substring(len2 - 8, 1);//1byte目7bit目
#if DEBUG//デバッグ時だけ実行
                        txtResponseDataByte_One_Bit0.Text = responseDataByte_One_Bit0;
                        txtResponseDataByte_One_Bit1.Text = responseDataByte_One_Bit1;
                        txtResponseDataByte_One_Bit2.Text = responseDataByte_One_Bit2;
                        txtResponseDataByte_One_Bit3.Text = responseDataByte_One_Bit3;
                        txtResponseDataByte_One_Bit4.Text = responseDataByte_One_Bit4;
                        txtResponseDataByte_One_Bit5.Text = responseDataByte_One_Bit5;
                        txtResponseDataByte_One_Bit6.Text = responseDataByte_One_Bit6;
                        txtResponseDataByte_One_Bit7.Text = responseDataByte_One_Bit7;

#endif//常に実行

                        //1byte目0bit目(給電ボタン)
                        if (responseDataByte_One_Bit0 == "1")
                        {
                            //現在状況ラベル更新
                            lblGenzaiJyokyo.Text = "給電";

                            //給電ボタン背景色(緑)
                            groupBox4.BackColor = Color.MediumSeaGreen;
                        }

                        //1byte目1bit目(充電ボタン)
                        if (responseDataByte_One_Bit1 == "1")
                        {
                            //現在状況ラベル更新
                            lblGenzaiJyokyo.Text = "充電";

                            //緑色背景に変更
                            groupBox4.BackColor = Color.FromArgb(255, 217, 102);
                        }

                        if (responseDataByte_One_Bit0 == "0" && responseDataByte_One_Bit1 == "0")
                        {
                            //白背景に変更
                            groupBox4.BackColor = Color.WhiteSmoke;
                            //現在状況ラベル更新
                            lblGenzaiJyokyo.Text = "OFF";
                        }

                        //1byte目2bit目(終了フラグ)
                        if (responseDataByte_One_Bit2 == "1")
                        {
                            resetBox.Image = Properties.Resources.丸ボタン_線幅10_青色01_100p;
                        }
                        else
                        {
                            resetBox.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }

                        //1byte目4bit目 警告状態
                        if (responseDataByte_One_Bit4 == "1")
                        {
                            keikokuBox.Image = Properties.Resources.丸ボタン_線幅10_赤色02_100p;
                        }
                        else
                        {
                            keikokuBox.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }

                        //1byte目6bit目 警告状態
                        if (responseDataByte_One_Bit6 == "1")
                        {
                            warningBox.Image = Properties.Resources.丸ボタン_線幅10_赤色02_100p;
                        }
                        else
                        {
                            warningBox.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }

                        //---------------------
                        //2byte目を２進数で抽出
                        //---------------------
                        string responseDataByte_Two = Convert.ToString(responseDataBytes[2], 2).PadLeft(8, '0');
                        txtResponseDataByte_Two.Text = responseDataByte_Two;

                        string responseDataByte_Two_Bit0 = responseDataByte_Two.Substring(len2 - 1, 1); //2byte目0bit目
                        string responseDataByte_Two_Bit1 = responseDataByte_Two.Substring(len2 - 2, 1); //2byte目1bit目
                        string responseDataByte_Two_Bit2 = responseDataByte_Two.Substring(len2 - 3, 1); //2byte目2bit目
                        string responseDataByte_Two_Bit3 = responseDataByte_Two.Substring(len2 - 4, 1); //2byte目3bit目
                        string responseDataByte_Two_Bit4 = responseDataByte_Two.Substring(len2 - 5, 1); //2byte目4bit目
                        string responseDataByte_Two_Bit5 = responseDataByte_Two.Substring(len2 - 6, 1); //2byte目5bit目
                        string responseDataByte_Two_Bit6 = responseDataByte_Two.Substring(len2 - 7, 1); //2byte目6bit目
                        string responseDataByte_Two_Bit7 = responseDataByte_Two.Substring(len2 - 8, 1); //2byte目7bit目
#if DEBUG//デバッグ時だけ実行
                        txtResponseDataByte_Two_Bit0.Text = responseDataByte_Two_Bit0;
                        txtResponseDataByte_Two_Bit1.Text = responseDataByte_Two_Bit1;
                        txtResponseDataByte_Two_Bit2.Text = responseDataByte_Two_Bit2;
                        txtResponseDataByte_Two_Bit3.Text = responseDataByte_Two_Bit3;
                        txtResponseDataByte_Two_Bit4.Text = responseDataByte_Two_Bit4;
                        txtResponseDataByte_Two_Bit5.Text = responseDataByte_Two_Bit5;
                        txtResponseDataByte_Two_Bit6.Text = responseDataByte_Two_Bit6;
                        txtResponseDataByte_Two_Bit7.Text = responseDataByte_Two_Bit7;

#endif//常に実行

                        //2byte目0bit目(R1)の絞り込み
                        if (responseDataByte_Two_Bit0 == "0")
                        {
                            R1.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }
                        else
                        {
                            R1.Image = Properties.Resources.丸ボタン_線幅10_緑色03_100p;
                        }

                        //2byte目1bit目(R2)の絞り込み
                        if (responseDataByte_Two_Bit1 == "0")
                        {
                            R2.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }
                        else
                        {
                            R2.Image = Properties.Resources.丸ボタン_線幅10_緑色03_100p;
                        }

                        //2byte目2bit目(R3)の絞り込み
                        if (responseDataByte_Two_Bit2 == "0")
                        {
                            R3.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }
                        else
                        {
                            R3.Image = Properties.Resources.丸ボタン_線幅10_緑色03_100p;
                        }

                        //2byte目3bit目(R4)の絞り込み
                        if (responseDataByte_Two_Bit3 == "0")
                        {
                            R4.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }
                        else
                        {
                            R4.Image = Properties.Resources.丸ボタン_線幅10_緑色03_100p;
                        }

                        //2byte目4bit目(R5)の絞り込み
                        if (responseDataByte_Two_Bit4 == "0")
                        {
                            R5.Image = Properties.Resources.丸ボタン_線幅10_灰色_100p;
                        }
                        else
                        {
                            R5.Image = Properties.Resources.丸ボタン_線幅10_緑色03_100p;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        //受信イベント01
        private void serialPort1_DataReceived_1(object sender, SerialDataReceivedEventArgs e)
        {
            string str = serialPort1.ReadLine();
            Response(str);
        }


        //給電ボタンイベント
        private void kyuudenbutton_Click(object sender, EventArgs e)
        {
            //給電OFF
            if (kyuudenbuttonFlg)
            {
                kyuudenbuttonFlg = false;
                ////textBox1.AppendText("kyuudenbuttonFlg = false\r\n");
                try
                {
#if DEBUG//デバッグ時だけ実行
                    textBox1.AppendText("給電OFFデータを送信しました\r\n");
#endif//常に実行
                    //変更するデータ
                    sendByte[1] = 0x0a;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Porttoji();
                }
            }
            //給電ON
            else
            {
                kyuudenbuttonFlg = true;
                ////textBox1.AppendText("kyuudenbuttonFlg = true\r\n");
                try
                {
#if DEBUG//デバッグ時だけ実行
                    textBox1.AppendText("給電ONデータを送信しました\r\n");
#endif//常に実行
                    //変更するデータ
                    sendByte[1] = 0x01;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Porttoji();
                }
            }

            botanhantei();//充電ボタン・給電ボタン判定
        }
        //充電ボタンイベント
        private void juudenbutton_Click(object sender, EventArgs e)
        {
            //充電OFF
            if (judenbuttonFlg)
            {
#if DEBUG//デバッグ時だけ実行
                textBox1.AppendText("judenbuttonFlg = false\r\n");
                textBox1.AppendText("充電OFFデータを送信しました\r\n");
#endif//常に実行
                judenbuttonFlg = false;
                //変更するデータ
                sendByte[1] = 0x08;
            }
            //充電OFF
            else
            {
#if DEBUG//デバッグ時だけ実行
                textBox1.AppendText("judenbuttonFlg = true\r\n");
                textBox1.AppendText("充電ONデータを送信しました\r\n");
#endif//常に実行
                judenbuttonFlg = true;
                //変更するデータ
                sendByte[1] = 0x04;
            }
            botanhantei();//充電ボタン・給電ボタン判定
        }

        //充電ボタン・給電ボタン判定
        //緑MediumSeaGreen
        //黄色255,217,102
        private void botanhantei()
        {
            //給電ボタンが押されていたら
            if (kyuudenbuttonFlg)
            {
                juudenbutton.Enabled = false;    // 押せない状態に
                //label3.Text = "給電";
                //緑色背景に変更
                //groupBox4.BackColor = Color.MediumSeaGreen;
                juudenbutton.BackColor = Color.Gray;
            }
            //デフォルト
            else
            {
                juudenbutton.Enabled = true;     // 押せる状態に
                juudenbutton.BackColor = Color.FromArgb(255, 217, 102);
            }
            //充電ボタンが押されていたら
            if (judenbuttonFlg)
            {
                kyuudenbutton.Enabled = false;    // 押せない状態に
                //label3.Text = "充電";
                //黄色背景に変更
                //groupBox4.BackColor = Color.FromArgb(255, 217, 102);
                juudenbutton.BackColor = Color.FromArgb(255, 217, 102);
                kyuudenbutton.BackColor = Color.Gray;
            }
            else
            {
                kyuudenbutton.Enabled = true;     // 押せる状態に
                kyuudenbutton.BackColor = Color.MediumSeaGreen;
            }
        }

        //異常模擬ボタンクリックイベント
        private void ijoumogibutton_Click(object sender, EventArgs e)
        {
            //異常模擬OFF
            if (ijoumogibuttonFlg)
            {
#if DEBUG//デバッグ時だけ実行
                textBox1.AppendText("ijoumogibuttonFlg = false\r\n異常模擬OFFデータを送信しました\r\n");
#endif//常に実行
                ijoumogibutton.ForeColor = Color.Black;
                ijoumogibutton.BackColor = Color.LightCoral;
                ijoumogibuttonFlg = false;
                //変更するデータ
                sendByte[2] = 0x00;
            }
            //異常模擬ON
            else
            {
#if DEBUG//デバッグ時だけ実行
                textBox1.AppendText("ijoumogibuttonFlg = true\r\n異常模擬ONデータを送信しました\r\n");
#endif//常に実行
                ijoumogibutton.ForeColor = Color.White;
                ijoumogibutton.BackColor = Color.Firebrick;
                ijoumogibuttonFlg = true;
                //変更するデータ
                sendByte[2] = 0x01;
            }
        }

        //送信ボタン設定
        private void sousinbutton_Click(object sender, EventArgs e)
        {
#if DEBUG//デバッグ時だけ実行
#endif//常に実行
        }
        //切断ボタンイベント
        private void setudanbutton_Click(object sender, EventArgs e)
        {
            //! シリアルポートをオープンしている場合、クローズする.
            if (serialPort1.IsOpen == true)
            {
                try
                {
                    serialPort1.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                if (serialPort1.IsOpen == false)
                {
#if DEBUG//デバッグ時だけ実行
                    textBox1.AppendText("切断しました。\r\n");
#endif//常に実行
                }
                Setuzokukousin();
            }
        }

        //接続設定ボタンの更新
        private void Setuzokukousin()
        {
            if (serialPort1.IsOpen == true)
            {
                //ボタンの表示を[切断]から[接続]に変える.
                setuzokubutton.Text = "切断";
            }
            if (serialPort1.IsOpen == false)
            {
                //ボタンの表示を[接続]から[切断]に変える.
                setuzokubutton.Text = "接続";
            }
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                label2.Text = (port);
            }
        }

        //リセットボタン
        private void resetbutton_Click(object sender, EventArgs e)
        {
            //変更するデータ
            sendByte[2] = 0x64;

            textBox1.AppendText("リセットボタンイベント発生\r\n");
            //繰り返し処理
            for (int xx = 1; xx < 4; xx++)
            {
                //変更するバイトデータを書き込む
                serialPort1.Write(sendByte, 0, sendByte.Length);

                textBox1.AppendText(xx + "回目のループです\r\n");
            }
            //変更するデータを戻す
            sendByte[2] = 0x00;
        }

        //ポートが開いていたら閉じるイベント
        private void Porttoji()
        {
            //もしポートが開いていたら
            if (serialPort1.IsOpen == true)
            {
                //シリアルポートを閉じる
                serialPort1.Close();
            }
        }
        //タイマー繰り返し処理
        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                //変更するバイトデータを書き込む
                serialPort1.Write(sendByte, 0, sendByte.Length);
            }
            catch (Exception ex)
            {
                //タイマーを停止する
                timer1.Stop();

                //シリアルポートを閉じる
                serialPort1.Close();

                //ボタンの表示を[切断]から[接続]に変える
                setuzokubutton.Text = "接続";

                //コンボボックスの外部からの入力 true:有効 false:無効化
                comboBox1.Enabled = true;

                //エラーメッセージ
                MessageBox.Show(ex.Message);
            }
        }
        //タイマー開始・停止処理
        private void Timersettei()
        {
            //ポートが開いていたら
            if (serialPort1.IsOpen == true)
            {
                //タイマーを開始する
                timer1.Start();
            }
            //ポートが閉じていたら
            if (serialPort1.IsOpen == false)
            {
                //タイマーを停止する
                timer1.Stop();
            }
        }
    }
}
