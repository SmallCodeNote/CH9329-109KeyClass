using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Ports;

namespace CH9329
{
    public class CH9329
    {
        public string PortName;
        public int BaudRate;
        public int xSize;
        public int ySize;

        SerialPort serialPort;

        public Queue<string> MessageLog= new Queue<string>();

        public int MessageLogCount = 32;
        private void addMessageLog(string message)
        {
            if (MessageLog.Count > MessageLogCount) { MessageLog.Dequeue(); };
            MessageLog.Enqueue(message);
        }

        public string getMessageLog()
        {
            return String.Join("\r\n", MessageLog);
        }


        public CH9329(string PortName = "COM5", int xSize = 1920, int ySize = 1080, int BaudRate = 9600)
        {
            this.PortName = PortName;
            this.BaudRate = BaudRate;
            this.xSize = xSize;
            this.ySize = ySize;

            serialPort = new SerialPort(PortName, BaudRate);

            serialPort.Open();
            createCharKeyTable();
            createMediaKeyTable();
            createKeyTable();

        }

        private Dictionary<mediaKey, byte[]> mediaKeyTable;

        public enum mediaKey
        {
            EJECT,
            CDSTOP,
            PREVTRACK,
            NEXTTRACK,
            PLAYPAUSE,
            MUTE,
            VOLUMEDOWN,
            VOLUMEUP,
        }

        private void createMediaKeyTable()
        {
            mediaKeyTable = new Dictionary<mediaKey, byte[]>();

            mediaKeyTable.Add(mediaKey.EJECT, new byte[] { 0x02, 0x80, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.CDSTOP, new byte[] { 0x02, 0x40, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.PREVTRACK, new byte[] { 0x02, 0x20, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.NEXTTRACK, new byte[] { 0x02, 0x10, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.PLAYPAUSE, new byte[] { 0x02, 0x08, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.MUTE, new byte[] { 0x02, 0x04, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.VOLUMEDOWN, new byte[] { 0x02, 0x02, 0x00, 0x00 });
            mediaKeyTable.Add(mediaKey.VOLUMEUP, new byte[] { 0x02, 0x01, 0x00, 0x00 });

        }

        public Dictionary<byte[], string> keyTable;

        /// <summary>
        /// create 109A KeyTable
        /// </summary>
        private void createKeyTable()
        {
            keyTable = new Dictionary<byte[], string>();

            keyTable.Add(new byte[] { 0x00, 0x04 }, "A");   //001
            keyTable.Add(new byte[] { 0x00, 0x05 }, "B");   //002
            keyTable.Add(new byte[] { 0x00, 0x06 }, "C");   //003
            keyTable.Add(new byte[] { 0x00, 0x07 }, "D");   //004
            keyTable.Add(new byte[] { 0x00, 0x08 }, "E");   //005
            keyTable.Add(new byte[] { 0x00, 0x09 }, "F");   //006
            keyTable.Add(new byte[] { 0x00, 0x0A }, "G");   //007
            keyTable.Add(new byte[] { 0x00, 0x0B }, "H");   //008
            keyTable.Add(new byte[] { 0x00, 0x0C }, "I");   //009
            keyTable.Add(new byte[] { 0x00, 0x0D }, "J");   //010
            keyTable.Add(new byte[] { 0x00, 0x0E }, "K");   //011
            keyTable.Add(new byte[] { 0x00, 0x0F }, "L");   //012
            keyTable.Add(new byte[] { 0x00, 0x10 }, "M");   //013
            keyTable.Add(new byte[] { 0x00, 0x11 }, "N");   //014
            keyTable.Add(new byte[] { 0x00, 0x12 }, "O");   //015
            keyTable.Add(new byte[] { 0x00, 0x13 }, "P");   //016
            keyTable.Add(new byte[] { 0x00, 0x14 }, "Q");   //017
            keyTable.Add(new byte[] { 0x00, 0x15 }, "R");   //018
            keyTable.Add(new byte[] { 0x00, 0x16 }, "S");   //019
            keyTable.Add(new byte[] { 0x00, 0x17 }, "T");   //020
            keyTable.Add(new byte[] { 0x00, 0x18 }, "U");   //021
            keyTable.Add(new byte[] { 0x00, 0x19 }, "V");   //022
            keyTable.Add(new byte[] { 0x00, 0x1A }, "W");   //023
            keyTable.Add(new byte[] { 0x00, 0x1B }, "X");   //024
            keyTable.Add(new byte[] { 0x00, 0x1C }, "Y");   //025
            keyTable.Add(new byte[] { 0x00, 0x1D }, "Z");   //026
            keyTable.Add(new byte[] { 0x00, 0x1E }, "1");   //027
            keyTable.Add(new byte[] { 0x00, 0x1F }, "2");   //028
            keyTable.Add(new byte[] { 0x00, 0x20 }, "3");   //029
            keyTable.Add(new byte[] { 0x00, 0x21 }, "4");   //030
            keyTable.Add(new byte[] { 0x00, 0x22 }, "5");   //031
            keyTable.Add(new byte[] { 0x00, 0x23 }, "6");   //032
            keyTable.Add(new byte[] { 0x00, 0x24 }, "7");   //033
            keyTable.Add(new byte[] { 0x00, 0x25 }, "8");   //034
            keyTable.Add(new byte[] { 0x00, 0x26 }, "9");   //035
            keyTable.Add(new byte[] { 0x00, 0x27 }, "0");   //036
            keyTable.Add(new byte[] { 0x00, 0x28 }, "Enter");   //037
            keyTable.Add(new byte[] { 0x00, 0x29 }, "Esc"); //038
            keyTable.Add(new byte[] { 0x00, 0x2A }, "Backspace");   //039
            keyTable.Add(new byte[] { 0x00, 0x2B }, "Tab"); //040
            keyTable.Add(new byte[] { 0x00, 0x2C }, "Spacebar");    //041
            keyTable.Add(new byte[] { 0x00, 0x2D }, "-");   //042
            keyTable.Add(new byte[] { 0x00, 0x2E }, "^");   //043
            keyTable.Add(new byte[] { 0x00, 0x2F }, "@");   //044
            keyTable.Add(new byte[] { 0x00, 0x30 }, "[");   //045
            keyTable.Add(new byte[] { 0x00, 0x31 }, "-----");   //046
            keyTable.Add(new byte[] { 0x00, 0x32 }, "]");   //047
            keyTable.Add(new byte[] { 0x00, 0x33 }, ";");   //048
            keyTable.Add(new byte[] { 0x00, 0x34 }, ":");   //049
            keyTable.Add(new byte[] { 0x00, 0x35 }, "半角/全角");   //050
            keyTable.Add(new byte[] { 0x00, 0x36 }, ",");   //051
            keyTable.Add(new byte[] { 0x00, 0x37 }, ".");   //052
            keyTable.Add(new byte[] { 0x00, 0x38 }, "/");   //053
            keyTable.Add(new byte[] { 0x00, 0x39 }, "Caps Lock");    //054
            keyTable.Add(new byte[] { 0x00, 0x3A }, "F1");  //055
            keyTable.Add(new byte[] { 0x00, 0x3B }, "F2");  //056
            keyTable.Add(new byte[] { 0x00, 0x3C }, "F3");  //057
            keyTable.Add(new byte[] { 0x00, 0x3D }, "F4");  //058
            keyTable.Add(new byte[] { 0x00, 0x3E }, "F5");  //059
            keyTable.Add(new byte[] { 0x00, 0x3F }, "F6");  //060
            keyTable.Add(new byte[] { 0x00, 0x40 }, "F7");  //061
            keyTable.Add(new byte[] { 0x00, 0x41 }, "F8");  //062
            keyTable.Add(new byte[] { 0x00, 0x42 }, "F9");  //063
            keyTable.Add(new byte[] { 0x00, 0x43 }, "F10"); //064
            keyTable.Add(new byte[] { 0x00, 0x44 }, "F11"); //065
            keyTable.Add(new byte[] { 0x00, 0x45 }, "F12"); //066
            keyTable.Add(new byte[] { 0x00, 0x46 }, "Print Screen");    //067
            keyTable.Add(new byte[] { 0x00, 0x47 }, "Scroll Lock"); //068
            keyTable.Add(new byte[] { 0x00, 0x48 }, "Pause");   //069
            keyTable.Add(new byte[] { 0x00, 0x49 }, "Insert");  //070
            keyTable.Add(new byte[] { 0x00, 0x4A }, "Home");    //071
            keyTable.Add(new byte[] { 0x00, 0x4B }, "Page Up"); //072
            keyTable.Add(new byte[] { 0x00, 0x4C }, "Delete");  //073
            keyTable.Add(new byte[] { 0x00, 0x4D }, "End"); //074
            keyTable.Add(new byte[] { 0x00, 0x4E }, "Page Down");   //075
            keyTable.Add(new byte[] { 0x00, 0x4F }, "→");   //076
            keyTable.Add(new byte[] { 0x00, 0x50 }, "←");   //077
            keyTable.Add(new byte[] { 0x00, 0x51 }, "↓");   //078
            keyTable.Add(new byte[] { 0x00, 0x52 }, "↑");   //079
            keyTable.Add(new byte[] { 0x00, 0x53 }, "Num Lock");    //080
            keyTable.Add(new byte[] { 0x00, 0x54 }, "Keypad /");    //081
            keyTable.Add(new byte[] { 0x00, 0x55 }, "Keypad *");    //082
            keyTable.Add(new byte[] { 0x00, 0x56 }, "Keypad -");    //083
            keyTable.Add(new byte[] { 0x00, 0x57 }, "Keypad +");    //084
            keyTable.Add(new byte[] { 0x00, 0x58 }, "Keypad Enter");    //085
            keyTable.Add(new byte[] { 0x00, 0x59 }, "Keypad 1");    //086
            keyTable.Add(new byte[] { 0x00, 0x5A }, "Keypad 2");    //087
            keyTable.Add(new byte[] { 0x00, 0x5B }, "Keypad 3");    //088
            keyTable.Add(new byte[] { 0x00, 0x5C }, "Keypad 4");    //089
            keyTable.Add(new byte[] { 0x00, 0x5D }, "Keypad 5");    //090
            keyTable.Add(new byte[] { 0x00, 0x5E }, "Keypad 6");    //091
            keyTable.Add(new byte[] { 0x00, 0x5F }, "Keypad 7");    //092
            keyTable.Add(new byte[] { 0x00, 0x60 }, "Keypad 8");    //093
            keyTable.Add(new byte[] { 0x00, 0x61 }, "Keypad 9");    //094
            keyTable.Add(new byte[] { 0x00, 0x62 }, "Keypad 0");    //095
            keyTable.Add(new byte[] { 0x00, 0x63 }, "Keypad .");    //096
            keyTable.Add(new byte[] { 0x00, 0x65 }, "Application"); //097
            keyTable.Add(new byte[] { 0x00, 0x87 }, "\\");   //098
            keyTable.Add(new byte[] { 0x00, 0x88 }, "ひらがな カタカナ");   //099
            keyTable.Add(new byte[] { 0x00, 0x89 }, "\\");   //100
            keyTable.Add(new byte[] { 0x00, 0x8A }, "変換");  //101
            keyTable.Add(new byte[] { 0x00, 0x8B }, "無変換"); //102
            keyTable.Add(new byte[] { 0x00, 0xE0 }, "Left Ctrl");   //103
            keyTable.Add(new byte[] { 0x00, 0xE1 }, "Left Shift");  //104
            keyTable.Add(new byte[] { 0x00, 0xE2 }, "Left Alt");    //105
            keyTable.Add(new byte[] { 0x00, 0xE3 }, "Left Windows");    //106
            keyTable.Add(new byte[] { 0x00, 0xE4 }, "Right Ctrl");  //107
            keyTable.Add(new byte[] { 0x00, 0xE5 }, "Right Shift"); //108
            keyTable.Add(new byte[] { 0x00, 0xE6 }, "Right Alt");   //109
            keyTable.Add(new byte[] { 0x00, 0xE7 }, "Right Windows");   //110

            keyTable.Add(new byte[] { 0x02, 0x1E }, "!");   //111
            keyTable.Add(new byte[] { 0x02, 0x1F }, "\"");	//112
            keyTable.Add(new byte[] { 0x02, 0x20 }, "#");   //113
            keyTable.Add(new byte[] { 0x02, 0x21 }, "$");   //114
            keyTable.Add(new byte[] { 0x02, 0x22 }, "%");   //115
            keyTable.Add(new byte[] { 0x02, 0x23 }, "&");   //116
            keyTable.Add(new byte[] { 0x02, 0x24 }, "'");   //117
            keyTable.Add(new byte[] { 0x02, 0x25 }, "(");   //118
            keyTable.Add(new byte[] { 0x02, 0x26 }, ")");   //119
            keyTable.Add(new byte[] { 0x02, 0x2D }, "=");   //120
            keyTable.Add(new byte[] { 0x02, 0x2E }, "~");   //121
            keyTable.Add(new byte[] { 0x02, 0x2F }, "`");   //122
            keyTable.Add(new byte[] { 0x02, 0x30 }, "{");   //123
            keyTable.Add(new byte[] { 0x02, 0x32 }, "}");   //124
            keyTable.Add(new byte[] { 0x02, 0x33 }, "+");   //125
            keyTable.Add(new byte[] { 0x02, 0x34 }, "*");   //126
            keyTable.Add(new byte[] { 0x02, 0x36 }, "<");   //127
            keyTable.Add(new byte[] { 0x02, 0x37 }, ">");   //128
            keyTable.Add(new byte[] { 0x02, 0x38 }, "?");   //129
            keyTable.Add(new byte[] { 0x00, 0x87 }, "_");   //130
            keyTable.Add(new byte[] { 0x00, 0x89 }, "｜");	//131


        }

        private Dictionary<string, byte[]> charKeyTable;

        /// <summary>
        /// create 109A CharKeyTable
        /// </summary>
        private void createCharKeyTable()
        {
            charKeyTable = new Dictionary<string, byte[]>();
            
            charKeyTable.Add("0", new byte[] { 0, (byte)(0x27) });
            for (byte i = 1; i <= 9; i++)
            {
                charKeyTable.Add(i.ToString(), new byte[] { 0x00, (byte)(0x1E + i-1) });
            }

            for (byte i = 0; i < 26; i++)
            {
                //Upper ASCII code
                charKeyTable.Add(((char)(i + 65)).ToString(), new byte[] { 0x02, (byte)(0x04 + i) });
                //Lower ASCII code
                charKeyTable.Add(((char)(i + 97)).ToString(), new byte[] { 0x00, (byte)(0x04 + i) });
            }

            charKeyTable.Add("ENTER", new byte[] { 0x00, 0x28 });   //001
            charKeyTable.Add("ESC", new byte[] { 0x00, 0x29 }); //002
            charKeyTable.Add("BACKSPACE", new byte[] { 0x00, 0x2A });   //003
            charKeyTable.Add("TAB", new byte[] { 0x00, 0x2B }); //004
            charKeyTable.Add("SPACEBAR", new byte[] { 0x00, 0x2C });    //005
            charKeyTable.Add("-", new byte[] { 0x00, 0x2D });   //006
            charKeyTable.Add("^", new byte[] { 0x00, 0x2E });   //007
            charKeyTable.Add("@", new byte[] { 0x00, 0x2F });   //008
            charKeyTable.Add("[", new byte[] { 0x00, 0x30 });   //009
            charKeyTable.Add("-----", new byte[] { 0x00, 0x31 });   //010
            charKeyTable.Add("]", new byte[] { 0x00, 0x32 });   //011
            charKeyTable.Add(";", new byte[] { 0x00, 0x33 });   //012
            charKeyTable.Add(":", new byte[] { 0x00, 0x34 });   //013
            charKeyTable.Add("半角/全角", new byte[] { 0x00, 0x35 });   //014
            charKeyTable.Add(",", new byte[] { 0x00, 0x36 });   //015
            charKeyTable.Add(".", new byte[] { 0x00, 0x37 });   //016
            charKeyTable.Add("/", new byte[] { 0x00, 0x38 });   //017
            charKeyTable.Add("CAPS LOCK", new byte[] { 0x00, 0x39 });    //018
            charKeyTable.Add("F1", new byte[] { 0x00, 0x3A });  //019
            charKeyTable.Add("F2", new byte[] { 0x00, 0x3B });  //020
            charKeyTable.Add("F3", new byte[] { 0x00, 0x3C });  //021
            charKeyTable.Add("F4", new byte[] { 0x00, 0x3D });  //022
            charKeyTable.Add("F5", new byte[] { 0x00, 0x3E });  //023
            charKeyTable.Add("F6", new byte[] { 0x00, 0x3F });  //024
            charKeyTable.Add("F7", new byte[] { 0x00, 0x40 });  //025
            charKeyTable.Add("F8", new byte[] { 0x00, 0x41 });  //026
            charKeyTable.Add("F9", new byte[] { 0x00, 0x42 });  //027
            charKeyTable.Add("F10", new byte[] { 0x00, 0x43 }); //028
            charKeyTable.Add("F11", new byte[] { 0x00, 0x44 }); //029
            charKeyTable.Add("F12", new byte[] { 0x00, 0x45 }); //030
            charKeyTable.Add("PRINT SCREEN", new byte[] { 0x00, 0x46 });    //031
            charKeyTable.Add("SCROLL LOCK", new byte[] { 0x00, 0x47 }); //032
            charKeyTable.Add("PAUSE", new byte[] { 0x00, 0x48 });   //033
            charKeyTable.Add("INSERT", new byte[] { 0x00, 0x49 });  //034
            charKeyTable.Add("HOME", new byte[] { 0x00, 0x4A });    //035
            charKeyTable.Add("PAGE UP", new byte[] { 0x00, 0x4B }); //036
            charKeyTable.Add("DELETE", new byte[] { 0x00, 0x4C });  //037
            charKeyTable.Add("END", new byte[] { 0x00, 0x4D }); //038
            charKeyTable.Add("PAGE DOWN", new byte[] { 0x00, 0x4E });   //039
            charKeyTable.Add("→", new byte[] { 0x00, 0x4F });   //040
            charKeyTable.Add("←", new byte[] { 0x00, 0x50 });   //041
            charKeyTable.Add("↓", new byte[] { 0x00, 0x51 });   //042
            charKeyTable.Add("↑", new byte[] { 0x00, 0x52 });   //043
            charKeyTable.Add("NUM LOCK", new byte[] { 0x00, 0x53 });    //044
            charKeyTable.Add("KEYPAD /", new byte[] { 0x00, 0x54 });    //045
            charKeyTable.Add("KEYPAD *", new byte[] { 0x00, 0x55 });    //046
            charKeyTable.Add("KEYPAD -", new byte[] { 0x00, 0x56 });    //047
            charKeyTable.Add("KEYPAD +", new byte[] { 0x00, 0x57 });    //048
            charKeyTable.Add("KEYPAD ENTER", new byte[] { 0x00, 0x58 });    //049
            charKeyTable.Add("KEYPAD 1", new byte[] { 0x00, 0x59 });    //050
            charKeyTable.Add("KEYPAD 2", new byte[] { 0x00, 0x5A });    //051
            charKeyTable.Add("KEYPAD 3", new byte[] { 0x00, 0x5B });    //052
            charKeyTable.Add("KEYPAD 4", new byte[] { 0x00, 0x5C });    //053
            charKeyTable.Add("KEYPAD 5", new byte[] { 0x00, 0x5D });    //054
            charKeyTable.Add("KEYPAD 6", new byte[] { 0x00, 0x5E });    //055
            charKeyTable.Add("KEYPAD 7", new byte[] { 0x00, 0x5F });    //056
            charKeyTable.Add("KEYPAD 8", new byte[] { 0x00, 0x60 });    //057
            charKeyTable.Add("KEYPAD 9", new byte[] { 0x00, 0x61 });    //058
            charKeyTable.Add("KEYPAD 0", new byte[] { 0x00, 0x62 });    //059
            charKeyTable.Add("KEYPAD .", new byte[] { 0x00, 0x63 });    //060
            charKeyTable.Add("APPLICATION", new byte[] { 0x00, 0x65 }); //061
            //charKeyTable.Add("＼", new byte[] { 0x00, 0x87 });   //062
            charKeyTable.Add("ひらがな カタカナ", new byte[] { 0x00, 0x88 });   //063
            charKeyTable.Add("\\", new byte[] { 0x00, 0x89 });   //064
            charKeyTable.Add("変換", new byte[] { 0x00, 0x8A });  //065
            charKeyTable.Add("無変換", new byte[] { 0x00, 0x8B }); //066
            charKeyTable.Add("LEFT CTRL", new byte[] { 0x00, 0xE0 });   //067
            charKeyTable.Add("LEFT SHIFT", new byte[] { 0x00, 0xE1 });  //068
            charKeyTable.Add("LEFT ALT", new byte[] { 0x00, 0xE2 });    //069
            charKeyTable.Add("LEFT WINDOWS", new byte[] { 0x00, 0xE3 });    //070
            charKeyTable.Add("RIGHT CTRL", new byte[] { 0x00, 0xE4 });  //071
            charKeyTable.Add("RIGHT SHIFT", new byte[] { 0x00, 0xE5 }); //072
            charKeyTable.Add("RIGHT ALT", new byte[] { 0x00, 0xE6 });   //073
            charKeyTable.Add("RIGHT WINDOWS", new byte[] { 0x00, 0xE7 });   //074

            charKeyTable.Add("!", new byte[] { 0x02, 0x1E });   //075
            charKeyTable.Add("\"", new byte[] { 0x02, 0x1F });	//076
            charKeyTable.Add("#", new byte[] { 0x02, 0x20 });   //077
            charKeyTable.Add("$", new byte[] { 0x02, 0x21 });   //078
            charKeyTable.Add("%", new byte[] { 0x02, 0x22 });   //079
            charKeyTable.Add("&", new byte[] { 0x02, 0x23 });   //080
            charKeyTable.Add("'", new byte[] { 0x02, 0x24 });   //081
            charKeyTable.Add("(", new byte[] { 0x02, 0x25 });   //082
            charKeyTable.Add(")", new byte[] { 0x02, 0x26 });   //083
            charKeyTable.Add("=", new byte[] { 0x02, 0x2D });   //084
            charKeyTable.Add("~", new byte[] { 0x02, 0x2E });   //085
            charKeyTable.Add("`", new byte[] { 0x02, 0x2F });   //086
            charKeyTable.Add("{", new byte[] { 0x02, 0x30 });   //087
            charKeyTable.Add("}", new byte[] { 0x02, 0x32 });   //088
            charKeyTable.Add("+", new byte[] { 0x02, 0x33 });   //089
            charKeyTable.Add("*", new byte[] { 0x02, 0x34 });   //090
            charKeyTable.Add("<", new byte[] { 0x02, 0x36 });   //091
            charKeyTable.Add(">", new byte[] { 0x02, 0x37 });   //092
            charKeyTable.Add("?", new byte[] { 0x02, 0x38 });	//093

        }


        public enum SpecialKeyCode : byte
        {
            ENTER = 0x28,
            ESCAPE = 0x29,
            BACKSPACE = 0x2A,
            TAB = 0x2B,
            SPACEBAR = 0x2C,
            CAPS_LOCK = 0x39,
            F1 = 0x3A,
            F2 = 0x3B,
            F3 = 0x3C,
            F4 = 0x3D,
            F5 = 0x3E,
            F6 = 0x3F,
            F7 = 0x40,
            F8 = 0x41,
            F9 = 0x42,
            F10 = 0x43,
            F11 = 0x44,
            F12 = 0x45,
            PRINTSCREEN = 0x46,
            SCROLL_LOCK = 0x47,
            PAUSE = 0x48,
            INSERT = 0x49,
            HOME = 0x4A,
            PAGEUP = 0x4B,
            DELETE = 0x4C,
            END = 0x4D,
            PAGEDOWN = 0x4E,
            RIGHTARROW = 0x4F,
            LEFTARROW = 0x50,
            DOWNARROW = 0x51,
            UPARROW = 0x52,
            APPLICATION = 0x65,
            LEFT_CTRL = 0xE0,
            LEFT_SHIFT = 0xE1,
            LEFT_ALT = 0xE2,
            LEFT_WINDOWS = 0xE3,

            RIGHT_CTRL = 0xE4,
            RIGHT_SHIFT = 0xE5,
            RIGHT_ALT = 0xE6,
            RIGHT_WINDOWS = 0xE7,

            CTRL = 0xE4,
            SHIFT = 0xE5,
            ALT = 0xE6,
            WINDOWS = 0xE7,

        }

        public enum MouseButtonCode : byte
        {
            LEFT = 0x01,
            RIGHT = 0x02,
            MIDDLE = 0x04,
        }

        private string sendPacket(byte[] data)
        {
            serialPort.Write(data, 0, data.Length);
            Thread.Sleep(20);
            string resultMessage =  serialPort.ReadExisting();
            
            addMessageLog(data.ToString() + "|"+ resultMessage);
            return resultMessage;
        }

        private byte[] createPacketArray(List<int> arrList, bool addCheckSum)
        {
            List<byte> bytePacketList = arrList.ConvertAll(b => (byte)b);
            if (addCheckSum) bytePacketList.Add((byte)(arrList.Sum() & 0xff));
            return bytePacketList.ToArray();
        }

        /// <summary>
        /// charKeyUpPacket
        /// </summary>
        byte[] charKeyUpPacket = { 0x57, 0xAB, 0x00, 0x02, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0c };
        /// <summary>
        /// 
        /// mediaKeyUpPacket
        /// </summary>
        byte[] mediaKeyUpPacket = { 0x57, 0xAB, 0x00, 0x03, 0x04, 0x02, 0x00, 0x00, 0x00, 0x0B };

        /// <summary>
        /// KeyGroup
        /// </summary>
        public enum KeyGroup : byte
        {
            CharKey = 0x02,
            MediaKey = 0x03,
        }

        public enum CommandCode : byte
        {
            GET_INFO = 0x01,
            SEND_KB_GENERAL_DATA = 0x02,
            SEND_KB_MEDIA_DATA = 0x03,
            SEND_MS_ABS_DATA = 0x04,
            SEND_MS_REL_DATA = 0x05,
            READ_MY_HID_DATA = 0x07,
            GET_PARA_CFG = 0x08,
            GET_USB_STRING = 0x0A,
        }

        public byte CHIP_VERSION;
        public byte CHIP_STATUS;
        public bool NUM_LOCK;
        public bool CAPS_LOCK;
        public bool SCROLL_LOCK;

        public void getInfo()
        {
            byte[] getInfoPacket = { 0x57, 0xAB, 0x00, (byte)CommandCode.GET_INFO, 0x00,0x03 };
            string resultString = sendPacket(getInfoPacket);

            CHIP_VERSION = (byte)resultString[0];
            CHIP_STATUS =(byte)resultString[1];
            byte flagByte = (byte)resultString[2];
            NUM_LOCK = (flagByte & 0b00000001) > 0;
            CAPS_LOCK = (flagByte & 0b00000010) > 0;
            SCROLL_LOCK = (flagByte & 0b00000100) > 0;

        }

        /// <summary>
        /// Push key
        /// </summary>
        /// <param name="CMD">KetType</param>
        /// <param name="k0">special key code</param>
        /// <param name="k1">key code #1</param>
        /// <param name="k2">key code #2</param>
        /// <param name="k3">key code #3</param>
        /// <param name="k4">key code #4</param>
        /// <param name="k5">key code #5</param>
        /// <param name="k6">key code #6</param>
        public void keyDown(KeyGroup keyGroup,byte k0, byte k1, byte k2 = 0, byte k3 = 0, byte k4 = 0, byte k5 = 0, byte k6 = 0)
        {
            // ========================
            // keyDownPacketContents
            // HEAD{0x57, 0xAB} + ADDR{0x00} + CMD{0x02} + LEN{0x08} + DATA{k0, 0x00, k1, k2, k3, k4, k5, k6}
            // CMD = KeyGroup
            // ========================
            List<int> keyDownPacketListInt = new List<int> { 0x57, 0xAB, 0x00, (int)keyGroup, 0x08, k0, 0x00, k1, k2, k3, k4, k5, k6 };

            byte[] keyDownPacket = createPacketArray(keyDownPacketListInt, true);

            sendPacket(keyDownPacket);

        }
        public void keyUpAll()
        {
            keyUpAll(KeyGroup.CharKey);
        }
            public void keyUpAll(KeyGroup keyGroup)
        {
            if (keyGroup == KeyGroup.CharKey) { sendPacket(charKeyUpPacket); }
            else { sendPacket(mediaKeyUpPacket); };

        }
        public void keyDown(SpecialKeyCode specialKeyCode)
        {
            keyDown(KeyGroup.CharKey, (byte)specialKeyCode,0x00);

        }

        public void charKeyType(byte k0, byte k1, byte k2 = 0, byte k3 = 0, byte k4 = 0, byte k5 = 0, byte k6 = 0)
        {
            keyDown(KeyGroup.CharKey, k0, k1, k2, k3, k4, k5, k6);
            keyUpAll(KeyGroup.CharKey);
        }

        public void mediaKeyType(mediaKey MediaKey)
        {
            byte[] dat = mediaKeyTable[MediaKey];
            keyDown(KeyGroup.MediaKey, dat[0], dat[1], dat[2], dat[3]);
            keyUpAll(KeyGroup.MediaKey);

        }

        public void charKeyType(string typeString)
        {
            if (typeString.Length < 1) return;

            foreach (char s in typeString)
            {
                if (charKeyTable.ContainsKey(s.ToString()))
                {
                    byte[] dat = charKeyTable[s.ToString()];
                    charKeyType(dat[0], dat[1]);

                }

            }

        }

        public void mouseMoveAbs(int x, int y)
        {
            int xAbs = (int)(4096 * x / xSize);
            int yAbs = (int)(4096 * y / ySize);

            // ========================
            // mouseMoveAbsPacketContents
            // HEAD{0x57, 0xAB} + ADDR{0x00} + CMD{0x04} + LEN{0x07} + DATA{0x02, 0x00, [x],[x],[y],[y], 0x00}
            // CMD = 0x04 : USB mouse absolute mode
            // ========================
            List<int> mouseMoveAbsPacketListInt = new List<int> { 0x57, 0xAB, 0x00, 0x04, 0x07, 0x02, 0x00 };
            mouseMoveAbsPacketListInt.Add((byte)(xAbs & 0xff));
            mouseMoveAbsPacketListInt.Add((byte)(xAbs >> 8));
            mouseMoveAbsPacketListInt.Add((byte)(yAbs & 0xff));
            mouseMoveAbsPacketListInt.Add((byte)(yAbs >> 8));
            mouseMoveAbsPacketListInt.Add(0x00);

            byte[] mouseMoveAbsPacket = createPacketArray(mouseMoveAbsPacketListInt, true);
            sendPacket(mouseMoveAbsPacket);

        }


        public void mouseMoveRel(int x, int y)
        {
            if (x > 127) { x = 127; };
            if (x < -128) { x = -128; };
            if (x < 0) { x = 0x100 + x; };

            if (y > 127) { y = 127; };
            if (y < -128) { y = -128; };
            if (y < 0) { y = 0x100 + y; };

            // ========================
            // mouseMoveRelPacketContents
            // HEAD{0x57, 0xAB} + ADDR{0x00} + CMD{0x05} + LEN{0x05} + DATA{0x01, 0x00}
            // CMD = 0x05 : USB mouse relative mode
            // ========================
            List<int> mouseMoveRelPacketListInt = new List<int> { 0x57, 0xAB, 0x00, 0x05, 0x05, 0x01, 0x00 };
            mouseMoveRelPacketListInt.Add((byte)(x));
            mouseMoveRelPacketListInt.Add((byte)(y));
            mouseMoveRelPacketListInt.Add(0x00);

            byte[] mouseMoveRelPacket = createPacketArray(mouseMoveRelPacketListInt, true);
            sendPacket(mouseMoveRelPacket);

        }

        /// <summary>
        /// mouseButtonUpPacket
        /// </summary>
        byte[] mouseButtonUpPacket = { 0x57, 0xAB, 0x00, 0x05, 0x05, 0x01, 0x00, 0x00, 0x00, 0x00, 0x0D };

        public void mouseButtonDown(MouseButtonCode buttonCode)
        {
            // ========================
            // mouseClickPacketContents
            // HEAD{0x57, 0xAB} + ADDR{0x00} + CMD{0x05} + LEN{0x05} + DATA{0x01}
            // CMD = 0x05 : USB mouse relative mode
            // ========================
            List<int> mouseButtonDownPacketListInt = new List<int> { 0x57, 0xAB, 0x00, 0x05, 0x05, 0x01, 0x00, 0x00, 0x00, 0x00 };
            mouseButtonDownPacketListInt[6] = (int)buttonCode;

            byte[] mouseButtonDownPacket = createPacketArray(mouseButtonDownPacketListInt, true);
            sendPacket(mouseButtonDownPacket);
           
        }

        public void mouseButtonUpAll()
        {
            sendPacket(mouseButtonUpPacket);
        }

        public void mouseClick(MouseButtonCode buttonCode)
        {
            mouseButtonDown(buttonCode);
            mouseButtonUpAll();

        }

        public void mouseDoubleClick()
        {
            mouseClick(MouseButtonCode.LEFT);
            mouseClick(MouseButtonCode.LEFT);
        }

        public string mouseScroll(int scrollCount)
        {
            // ========================
            // mouseScrollPacketContents
            // HEAD{0x57, 0xAB} + ADDR{0x00} + CMD{0x05} + LEN{0x05} + DATA{0x01}
            // CMD = 0x05 : USB mouse relative mode
            // ========================
            if (scrollCount > 127) { scrollCount = 127; };
            if (scrollCount < -128) { scrollCount = -128; };
            if (scrollCount < 0) { scrollCount = 0x100 + scrollCount; };

            List<int> mouseScrollPacketListInt = new List<int> { 0x57, 0xAB, 0x00, 0x05, 0x05, 0x01, 0x00, 0x00, 0x00, 0x00 };
            mouseScrollPacketListInt.Add(scrollCount);

            byte[] mouseScrollPacket = createPacketArray(mouseScrollPacketListInt, true);
            return sendPacket(mouseScrollPacket);

        }

    }
}
