using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Management;
using CoreAudio;

namespace TrayOpen
{
    static class Program
    {
        #region Volume System
        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;
        private const int APPCOMMAND_MEDIA_PLAY_PAUSE = 0xE0000;


        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);
        #endregion


        static Hotkey OpenKey = new Hotkey(Keys.O, false, false, false, true);
        static Hotkey BlackKey = new Hotkey(Keys.Q, false, false, false, true);
        static Hotkey SleepKey = new Hotkey(Keys.S, false, false, false, true);
        static Hotkey HibernateKey = new Hotkey(Keys.H, false, false, false, true);
        static Hotkey VolumeDownKey = new Hotkey(Keys.OemOpenBrackets, false, false, false, true);
        static Hotkey VolumeUpKey = new Hotkey(Keys.OemCloseBrackets, false, false, false, true);


        static NotifyIcon Icon = new NotifyIcon();
        static ManagementEventWatcher w = null;
        static int WM_SYSCOMMAND = 0x0112;
        static int SC_MONITORPOWER = 0xF170;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Icon.Icon = new System.Drawing.Icon("ico.ico");
            ContextMenuStrip Menu = new ContextMenuStrip();
            Menu.Items.Add("Open Drive", null, delegate(object sender, EventArgs e) { Open(); });            
            Menu.Items.Add("Turn Off Screen", null, delegate(object sender, EventArgs e) { BlackScreen(); });
            Menu.Items.Add("Hibernate", null, delegate(object sender, EventArgs e) { Hibernate(); });
            Menu.Items.Add("Standby", null, delegate(object sender, EventArgs e) { Sleep(); });
            Menu.Items.Add("Exit", null, Exit);
            Icon.ContextMenuStrip = Menu;
            Icon.Visible = true;
            Icon.MouseDoubleClick += new MouseEventHandler(Icon_MouseDoubleClick);

            VolumeUpKey.Register(new Form());
            VolumeDownKey.Register(new Form());

            VolumeUpKey.Pressed += new System.ComponentModel.HandledEventHandler(VolumeUpKey_Pressed);
            VolumeDownKey.Pressed += new System.ComponentModel.HandledEventHandler(VolumeDownKey_Pressed);

            OpenKey.Register(new Form());           
            OpenKey.Pressed += new System.ComponentModel.HandledEventHandler(Key_Pressed);
            BlackKey.Register(new Form());
            BlackKey.Pressed += new System.ComponentModel.HandledEventHandler(Key_Pressed);
            HibernateKey.Register(new Form());
            HibernateKey.Pressed += new System.ComponentModel.HandledEventHandler(Key_Pressed);
            SleepKey.Register(new Form());                     
            SleepKey.Pressed += new System.ComponentModel.HandledEventHandler(Key_Pressed);
            ManagementEventWatcher EventWatcher = new ManagementEventWatcher();
            
            WqlEventQuery q;
            ManagementOperationObserver observer = new
            ManagementOperationObserver();

            // Bind to local machine
            ConnectionOptions opt = new ConnectionOptions();
            opt.EnablePrivileges = true; //sets required privilege
            ManagementScope scope = new ManagementScope("root\\CIMV2", opt);

            try
            {
                q = new WqlEventQuery();
                q.EventClassName = "__InstanceModificationEvent";
                q.WithinInterval = new TimeSpan(0, 0, 1);

                // DriveType - 5: CDROM
                q.Condition = @"TargetInstance ISA 'Win32_LogicalDisk' and TargetInstance.DriveType = 5";
                w = new ManagementEventWatcher(scope, q);

                // register async. event handler
                w.EventArrived += new EventArrivedEventHandler(CDREventArrived);
                w.Start();
            }
            catch { }

            System.IO.DriveInfo[] drives = System.IO.DriveInfo.GetDrives();
            foreach (System.IO.DriveInfo drive in drives)
            {
                if (drive.DriveType == System.IO.DriveType.CDRom)
                {
                    if (!drive.IsReady)
                    {
                        Icon.Visible = false;
                    }
                }
            }

            Application.Run();
        }

        static void VolumeDownKey_Pressed(object sender, System.ComponentModel.HandledEventArgs e)
        {
            //using (Form tmp = new Form())
            //{
            //    SendMessageW(tmp.Handle, WM_APPCOMMAND, tmp.Handle, (IntPtr)APPCOMMAND_VOLUME_DOWN);
            //}

            int i = PC_VolumeControl.VolumeControl.GetVolume();
            int v = i / 100;


            //MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            //MMDevice dev = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            //dev.AudioEndpointVolume.VolumeStepDown();
        }

        static void VolumeUpKey_Pressed(object sender, System.ComponentModel.HandledEventArgs e)
        {
            //using (Form tmp = new Form())
            //{
            //    SendMessageW(tmp.Handle, WM_APPCOMMAND, tmp.Handle, (IntPtr)APPCOMMAND_VOLUME_UP);
            //}

            PC_VolumeControl.VolumeControl.SetVolume(2);

            //MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            //MMDevice dev = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            //dev.AudioEndpointVolume.VolumeStepUp();
        }
       
        static void Icon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Open();
        }

        static void Key_Pressed(object sender, System.ComponentModel.HandledEventArgs e)
        {
            switch (((Hotkey)sender).KeyCode)
            {
                case Keys.H:
                    {
                        Hibernate();
                        break;
                    }
                case Keys.O:
                    {
                        Open();
                        break;
                    }
                case Keys.S:
                    {
                        Sleep();
                        break;
                    }
                case Keys.Q:
                    {
                        BlackScreen();
                        break;
                    }
            }
        }

        public static void BlackScreen()
        {
            using (Form Tmp = new Form())
            {
                SendMessage(Tmp.Handle, WM_SYSCOMMAND, SC_MONITORPOWER, 2);
            }
        }

        public static void Sleep()
        {
            Application.SetSuspendState(PowerState.Suspend, true, false);
        }

        public static void Hibernate()
        {
            Application.SetSuspendState(PowerState.Hibernate, true, false);
        }

        public static void Open()
        {
            mciSendString("set cdaudio door open", null, 0, IntPtr.Zero);
        }

        public static void Exit(object sender, EventArgs e)
        {
            Icon.Visible = false;
            w.Stop();
            OpenKey.Unregister();
            BlackKey.Unregister();
            HibernateKey.Unregister();
            SleepKey.Unregister();
            VolumeDownKey.Unregister();
            VolumeUpKey.Unregister();
            Refreash();
            Application.Exit();
        }
        public static void CDREventArrived(object sender, EventArrivedEventArgs e)
        {
            PropertyData pd;
            if ((pd = e.NewEvent.Properties["TargetInstance"]) != null)
            {
                ManagementBaseObject mbo = pd.Value as ManagementBaseObject;
                // if CD removed VolumeName == null
                if (mbo.Properties["VolumeName"].Value != null)
                {
                    Icon.Text = mbo.Properties["VolumeName"].Value.ToString();
                    Icon.Visible = true;
                    Refreash();
                    return;
                }
            }
            Icon.Text = "No Disk";
            Icon.Visible = false;
            Refreash();
        }

        private static void Refreash()
        {
            IntPtr traynotifywnd = FindWindow("Shell_TrayWnd", null);
            int result = SendMessage(traynotifywnd, 0xf, 0, 0);
        }

        // DLL access
        [DllImport("winmm.dll", EntryPoint = "mciSendStringA", CharSet = CharSet.Ansi)]
        static extern int mciSendString(string lpstrCommand, StringBuilder lpstrReturnString, int uReturnLength, IntPtr hwndCallback);

        public const int WM_PAINT = 0xF;
        [DllImport("USER32.DLL")]
        public static extern int SendMessage(IntPtr hwnd, int msg, int character, int lpsText);

        [DllImport("user32", EntryPoint = "FindWindowEx")]
        public static extern int FindWindowExA(int hWnd1, int hWnd2, string lpsz1, string lpsz2);

        [DllImport("user32", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindow(string classname, string windowname);
    }
}
