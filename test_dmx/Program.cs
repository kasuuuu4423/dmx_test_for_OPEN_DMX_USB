using System;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;

namespace test_dmx
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enttec Open DMX USB コントローラー");
            Console.WriteLine("================================");

            // 利用可能なシリアルポートを表示
            string[] ports = SerialPort.GetPortNames();
            if (ports.Length == 0)
            {
                Console.WriteLine("シリアルポートが見つかりません。");
                Console.WriteLine("Enttec Open DMX USBが接続されていることを確認してください。");
                Console.WriteLine("何かキーを押して終了...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("利用可能なシリアルポート:");
            for (int i = 0; i < ports.Length; i++)
            {
                Console.WriteLine($"{i + 1}: {ports[i]}");
            }

            // ポート選択
            Console.Write("使用するポート番号を選択してください: ");
            int portIndex;
            while (!int.TryParse(Console.ReadLine(), out portIndex) || portIndex < 1 || portIndex > ports.Length)
            {
                Console.Write("無効な選択です。再度入力してください: ");
            }
            string selectedPort = ports[portIndex - 1];

            // DMXコントローラーの初期化
            DmxController controller = new DmxController(selectedPort);
            
            try
            {
                controller.Open();
                Console.WriteLine($"ポート {selectedPort} に接続しました。");
                
                bool running = true;
                while (running)
                {
                    Console.WriteLine("\nコマンド:");
                    Console.WriteLine("1: チャンネル値を設定");
                    Console.WriteLine("2: 全チャンネルをリセット");
                    Console.WriteLine("3: チェイスパターンを実行");
                    Console.WriteLine("4: 現在の値を表示");
                    Console.WriteLine("0: 終了");
                    Console.Write("選択: ");

                    if (int.TryParse(Console.ReadLine(), out int choice))
                    {
                        switch (choice)
                        {
                            case 0:
                                running = false;
                                break;
                            case 1:
                                SetChannelValue(controller);
                                break;
                            case 2:
                                controller.ResetAllChannels();
                                Console.WriteLine("全チャンネルをリセットしました。");
                                break;
                            case 3:
                                RunChasePattern(controller);
                                break;
                            case 4:
                                DisplayCurrentValues(controller);
                                break;
                            default:
                                Console.WriteLine("無効な選択です。");
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("無効な入力です。");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
            finally
            {
                controller.Close();
                Console.WriteLine("DMXコントローラーを閉じました。");
            }

            Console.WriteLine("プログラムを終了します。何かキーを押してください...");
            Console.ReadKey();
        }

        static void SetChannelValue(DmxController controller)
        {
            Console.Write("チャンネル番号 (1-512): ");
            if (!int.TryParse(Console.ReadLine(), out int channel) || channel < 1 || channel > 512)
            {
                Console.WriteLine("無効なチャンネル番号です。");
                return;
            }

            Console.Write("値 (0-255): ");
            if (!int.TryParse(Console.ReadLine(), out int value) || value < 0 || value > 255)
            {
                Console.WriteLine("無効な値です。");
                return;
            }

            controller.SetChannelValue(channel, (byte)value);
            Console.WriteLine($"チャンネル {channel} を {value} に設定しました。");
        }

        static void RunChasePattern(DmxController controller)
        {
            Console.Write("開始チャンネル (1-512): ");
            if (!int.TryParse(Console.ReadLine(), out int startChannel) || startChannel < 1 || startChannel > 512)
            {
                Console.WriteLine("無効なチャンネル番号です。");
                return;
            }

            Console.Write("チャンネル数 (1-512): ");
            if (!int.TryParse(Console.ReadLine(), out int channelCount) || channelCount < 1 || startChannel + channelCount - 1 > 512)
            {
                Console.WriteLine("無効なチャンネル数です。");
                return;
            }

            Console.Write("値 (0-255): ");
            if (!int.TryParse(Console.ReadLine(), out int value) || value < 0 || value > 255)
            {
                Console.WriteLine("無効な値です。");
                return;
            }

            Console.Write("ステップ間の遅延（ミリ秒）: ");
            if (!int.TryParse(Console.ReadLine(), out int delay) || delay < 0)
            {
                Console.WriteLine("無効な遅延値です。");
                return;
            }

            Console.WriteLine("チェイスパターンを実行中... (Ctrl+C で停止)");
            
            try
            {
                // 最初に全てのチャンネルをリセット
                controller.ResetAllChannels();
                
                // チェイスパターンを実行
                int currentChannel = startChannel;
                bool running = true;
                
                // キャンセル用のトークンソース
                CancellationTokenSource cts = new CancellationTokenSource();
                
                // コンソールキー入力を監視するタスク
                Task.Run(() => {
                    Console.WriteLine("何かキーを押すと停止します...");
                    Console.ReadKey(true);
                    cts.Cancel();
                });
                
                while (running && !cts.Token.IsCancellationRequested)
                {
                    // 前のチャンネルをオフ
                    if (currentChannel > startChannel)
                    {
                        controller.SetChannelValue(currentChannel - 1, 0);
                    }
                    else if (currentChannel == startChannel && channelCount > 1)
                    {
                        controller.SetChannelValue(startChannel + channelCount - 1, 0);
                    }
                    
                    // 現在のチャンネルをオン
                    controller.SetChannelValue(currentChannel, (byte)value);
                    
                    // 次のチャンネルに進む
                    currentChannel++;
                    if (currentChannel > startChannel + channelCount - 1)
                    {
                        currentChannel = startChannel;
                    }
                    
                    Thread.Sleep(delay);
                }
                
                // 終了時に全チャンネルをリセット
                controller.ResetAllChannels();
                Console.WriteLine("チェイスパターンを停止しました。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
        }

        static void DisplayCurrentValues(DmxController controller)
        {
            byte[] values = controller.GetAllChannelValues();
            Console.WriteLine("現在のDMX値:");
            
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] > 0)
                {
                    Console.WriteLine($"チャンネル {i + 1}: {values[i]}");
                }
            }
            
            Console.WriteLine("値が0のチャンネルは表示されていません。");
        }
    }

    public class DmxController
    {
        private SerialPort _serialPort;
        private byte[] _dmxData;
        private readonly object _lockObject = new object();
        private bool _isOpen = false;
        private Thread? _sendThread;
        private bool _keepSending = false;

        public DmxController(string portName)
        {
            _serialPort = new SerialPort
            {
                PortName = portName,
                BaudRate = 250000,  // DMXのボーレート
                DataBits = 8,
                Parity = Parity.None,
                StopBits = StopBits.Two,
                Handshake = Handshake.None
            };

            // DMXデータの初期化 (512チャンネル + スタートコード)
            _dmxData = new byte[513];
            _dmxData[0] = 0; // スタートコード
        }

        public void Open()
        {
            if (!_isOpen)
            {
                _serialPort.Open();
                _isOpen = true;
                
                // 送信スレッドの開始
                _keepSending = true;
                _sendThread = new Thread(SendDmxDataContinuously);
                _sendThread.IsBackground = true;
                _sendThread.Start();
            }
        }

        public void Close()
        {
            if (_isOpen)
            {
                _keepSending = false;
                if (_sendThread != null && _sendThread.IsAlive)
                {
                    _sendThread.Join(1000); // スレッドの終了を待つ
                }
                
                _serialPort.Close();
                _isOpen = false;
            }
        }

        public void SetChannelValue(int channel, byte value)
        {
            if (channel < 1 || channel > 512)
                throw new ArgumentOutOfRangeException(nameof(channel), "チャンネルは1から512の範囲である必要があります。");

            lock (_lockObject)
            {
                _dmxData[channel] = value;
            }
        }

        public byte GetChannelValue(int channel)
        {
            if (channel < 1 || channel > 512)
                throw new ArgumentOutOfRangeException(nameof(channel), "チャンネルは1から512の範囲である必要があります。");

            lock (_lockObject)
            {
                return _dmxData[channel];
            }
        }

        public byte[] GetAllChannelValues()
        {
            lock (_lockObject)
            {
                // スタートコードを除いたデータをコピー
                byte[] values = new byte[512];
                Array.Copy(_dmxData, 1, values, 0, 512);
                return values;
            }
        }

        public void ResetAllChannels()
        {
            lock (_lockObject)
            {
                // スタートコードは保持
                for (int i = 1; i < _dmxData.Length; i++)
                {
                    _dmxData[i] = 0;
                }
            }
        }

        private void SendDmxDataContinuously()
        {
            while (_keepSending && _isOpen)
            {
                try
                {
                    SendDmxData();
                    Thread.Sleep(25); // 40Hzで送信（DMX更新レート）
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"DMX送信エラー: {ex.Message}");
                    Thread.Sleep(1000); // エラー時は少し待機
                }
            }
        }

        private void SendDmxData()
        {
            if (!_isOpen)
                return;

            lock (_lockObject)
            {
                try
                {
                    // DMXブレーク信号の送信
                    _serialPort.BreakState = true;
                    Thread.Sleep(1); // ブレーク信号の長さ（最低88μs）
                    _serialPort.BreakState = false;
                    Thread.Sleep(1); // MAB（Mark After Break）の長さ（最低8μs）

                    // DMXデータの送信
                    _serialPort.Write(_dmxData, 0, _dmxData.Length);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"DMX送信中にエラーが発生しました: {ex.Message}");
                    throw;
                }
            }
        }
    }
}
