using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using IniParser;
using IniParser.Model;
using Microsoft.Win32;

namespace mypc_tools
{
    public partial class Form1 : Form
    {
        public static Point courseraLeftTop = new Point(442, 225);
        public static Size courseraSize = new Size(940, 585);        

        public static Point youTubeLeftTop = new Point(310, 160);
        public static Size youTubeSize = new Size(965, 620);
        public static String startupName = "mypc_tools";

        private string rootPath = "C:\\workspace\\mypc_tools_service\\";
        FileIniDataParser parser = new FileIniDataParser();
        IniData initData;
        KeyboardHook hook = new KeyboardHook();

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void keybd_event(uint bVk, uint bScan, uint dwFlags, uint dwExtraInfo);
        const int KEYUP = 0x0002;

        public Form1()
        {
            InitializeComponent();

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
            this.notifyIcon1.Visible = true;          
            notifyIcon1.ContextMenuStrip = contextMenuStrip1;

            rootPath = AppDomain.CurrentDomain.BaseDirectory;
            System.IO.Directory.SetCurrentDirectory(rootPath);
            initData = parser.ReadFile(rootPath + "config.ini");

            // register the event that is fired after the key press.
            hook.KeyPressed += new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);

            foreach (KeyData keyItem in initData["Shortcuts"])
            {
                string key = keyItem.KeyName;
                string value = keyItem.Value;
                bool hasShift = false;

                if (key.StartsWith("#") || key.Trim().Length == 0)
                    continue;

                if (key.Contains("Shift"))
                    hasShift = true;

                if (key.StartsWith("F"))
                {
                    ModifierKeys modifier = 0 + (hasShift ? mypc_tools.ModifierKeys.Shift : 0);
                    string[] arr = key.Split('+');
                    hook.RegisterHotKey(modifier, Keys.F1 + Int32.Parse(arr[0].Substring(1)) - 1);
                }
            }
        }

        void hook_KeyPressed(object sender, KeyPressedEventArgs e)
        {
            if (e.Key >= Keys.F1 && e.Key <= Keys.F12)
            {
                if (null != initData)
                {
                    string pressedKey = "F" + ((e.Key - Keys.F1) + 1).ToString();
                    if (e.Modifier == mypc_tools.ModifierKeys.Shift)
                    {
                        pressedKey += "+Shift";
                    }
                    simulateKey(initData["Shortcuts"][pressedKey]);
                }
            }
        }

        private void simulateKey(String text)
        {

            // get foreground window and type some text
            // https://stackoverflow.com/questions/115868/how-do-i-get-the-title-of-the-current-active-window-using-c
            IntPtr handle = GetForegroundWindow();

            // simulate key press
            // https://stackoverflow.com/questions/3047375/simulating-key-press-c-sharp
            keybd_event(0x10, 0, KEYUP, 0); // Shift Release
            SendKeys.SendWait(text);
            // PressKey(Keys.Return); // 가끔 엔터가 안됨. 특히 F3 (이건 Notepadd++에 단축키가 있어서 그런 듯. 다른 곳에서는 잘 됨)
        }

        public static void PressKey(Keys key)
        {
            PressKey(key, false);
            PressKey(key, true);
        }

        public static void PressKey(Keys key, bool up)
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            const int KEYEVENTF_KEYUP = 0x2;
            if (up)
            {
                keybd_event((byte)key, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
            }
            else
            {
                keybd_event((byte)key, 0x45, KEYEVENTF_EXTENDEDKEY, 0);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            
        }

        private void toggleService()
        {
            // https://stackoverflow.com/questions/16317378/how-to-stop-windows-service-programmatically

        }

        private void captureCourseraMenuItem_Click(object sender, EventArgs e)
        {
            captureScreen(courseraLeftTop.X, courseraLeftTop.Y, courseraSize.Width, courseraSize.Height);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.notifyIcon1.Visible = false;
            Application.Exit();
        }

        public void captureScreen(int x, int y, int width, int height)
        {
            // Capture
            //작업 표시줄을 제외한 영역 크기   
            int w = Screen.PrimaryScreen.WorkingArea.Width;
            int h = Screen.PrimaryScreen.WorkingArea.Height;

            // Determine the size of the "virtual screen", which includes all monitors.
            int screenLeft = SystemInformation.VirtualScreen.Left;
            int screenTop = SystemInformation.VirtualScreen.Top;
            int screenWidth = SystemInformation.VirtualScreen.Width;
            int screenHeight = SystemInformation.VirtualScreen.Height;

            //Bitmap 객체 생성   
            Bitmap bmp = new Bitmap(width, height);

            //Graphics 객체 생성   
            Graphics g = Graphics.FromImage(bmp);

            //Graphics 객체의 CopyFromScreen()메서드로 bitmap 객체에 Screen을 캡처하여 저장   

            //g.CopyFromScreen(screenLeft, screenTop, 0, 0, bmp.Size);
            g.CopyFromScreen(screenLeft + x, screenTop + y, 0, 0, new Size(width, height));

            Clipboard.SetImage((Image)bmp);
        }

        private void addStartupToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                // 시작프로그램 등록하는 레지스트리
                string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                RegistryKey strUpKey = Registry.LocalMachine.OpenSubKey(runKey);
                if (strUpKey.GetValue(startupName) == null)
                {
                    strUpKey.Close();
                    // https://milkoon1.tistory.com/30: System.Security.SecurityException > 보안 설정
                    // 컴퓨터\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32
                    strUpKey = Registry.LocalMachine.OpenSubKey(runKey, true);
                    // 시작프로그램 등록명과 exe경로를 레지스트리에 등록
                    strUpKey.SetValue(startupName, Application.ExecutablePath);
                }
                MessageBox.Show("Add Startup Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("레지스트리 등록 실패 > 관리자 권한으로 재시도해보세요.: " + ex);
            }
        }

        private void removeStartupToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                RegistryKey strUpKey = Registry.LocalMachine.OpenSubKey(runKey, true);
                // 레지스트리값 제거
                strUpKey.DeleteValue(startupName);
                MessageBox.Show("Remove Startup Success");
            }
            catch
            {
                MessageBox.Show("레지스트리 제거 실패 > 관리자 권한으로 재시도해보세요.");
            }
        }
    }
}
