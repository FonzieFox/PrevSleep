using System;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("PrevSleep App")]
[assembly: AssemblyVersion("1.0.3.5")]
[assembly: AssemblyProduct("PrevSleep")]
[assembly: AssemblyCompany("VixByte")]
[assembly: AssemblyDescription("Simple tray application to prevent sleep schedules on computers. Usefull on company provided hardware where you don't have access to the power configuration, but your work isn't intense enough to keep the system awake at all times.")]
[assembly: AssemblyFileVersion("1.0.3.5")]

class Program
{
    [Flags]
        public enum EXECUTION_STATE : uint
        {
            ES_AWAYMODE_REQUIRED = 0x00000040,
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
        }

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    static extern uint SetThreadExecutionState(EXECUTION_STATE esFlags);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    static readonly System.Threading.TimerCallback PrvSlpCb = _ => {
        SetThreadExecutionState(EXECUTION_STATE.ES_DISPLAY_REQUIRED | EXECUTION_STATE.ES_SYSTEM_REQUIRED | EXECUTION_STATE.ES_CONTINUOUS);
        Point pos = Cursor.Position;
        Cursor.Position = new Point(pos.X, pos.Y + 1);
        Cursor.Position = pos;
        keybd_event(0x91, 0, 0x0000, UIntPtr.Zero);
        keybd_event(0x91, 0, 0x0002, UIntPtr.Zero);
        keybd_event(0x91, 0, 0x0000, UIntPtr.Zero);
        keybd_event(0x91, 0, 0x0002, UIntPtr.Zero);
    };

    [STAThread]
    static void Main()
    {
        System.Threading.Timer logicTimer = null;

        var trayIcon = new NotifyIcon(){
            //Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
            Text = "PrevSleep app",
            Visible = true,
            ContextMenuStrip = new ContextMenuStrip()
        };
        trayIcon.ContextMenuStrip.Items.Add("Exit", null, (s, e) => OnExit(trayIcon, logicTimer));

        DialogResult result = MessageBox.Show(
            "Welcome to PrevSleep app!\n" +
            "Remember that disabling the default sleeping function increases risk of unauthorised access to the workstation!\n" +
            "Always ensure your device is secured and you are logged off when not using the device!\n" +
            "The application provider is not responsible for any breach resulting in leaving the workstation unsecured while using the application.\n\n"+
            "Press YES to accept the risk and take the responsibility of any unauthorised access to the device.\n" +
            "Press NO to reject the risk and close the application\n", 
            "prevSleep - Consent",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information
        );

        if(result != DialogResult.Yes){
            ExitWithMessage(trayIcon, "Consent rejected");
            return;
        }

        uint testresult = SetThreadExecutionState(EXECUTION_STATE.ES_DISPLAY_REQUIRED | EXECUTION_STATE.ES_SYSTEM_REQUIRED | EXECUTION_STATE.ES_CONTINUOUS);
        if(testresult == 0)
        {
            ExitWithMessage(trayIcon, "Error - Could not suspend sleep schedule!");
            return;
        }

        trayIcon.ShowBalloonTip(3000, "App is running in the background", "Sleep suspended!", ToolTipIcon.Info);
        
        System.Threading.Timer clearTipTimer = null;
        clearTipTimer = new System.Threading.Timer(state => {
            try { trayIcon.BalloonTipTitle = null; trayIcon.BalloonTipText = null; }
            catch {}
            try { ((System.Threading.Timer)state).Dispose(); } catch {}
        }, null, 3500, System.Threading.Timeout.Infinite);

        logicTimer = new System.Threading.Timer(PrvSlpCb, null, 0, 180000);

        Application.Run();

        try { if (logicTimer != null) logicTimer.Dispose(); } catch { }
        try { trayIcon.Visible = false; } catch {}
        try { trayIcon.Dispose(); } catch {}
        try { SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS); } catch {}
    }

    static void OnExit(NotifyIcon trayIcon, System.Threading.Timer logicTimer){
        try { if (logicTimer != null) logicTimer.Dispose(); } catch { }
        try { trayIcon.Visible = false; } catch {}
        try { trayIcon.Dispose(); } catch {}
        try { SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS); } catch {}
        Application.Exit();
    }

    static void ExitWithMessage(NotifyIcon trayIcon, string message){
        try { trayIcon.ShowBalloonTip(2000, message, "Application is closing", ToolTipIcon.Error); } catch {}
        
        try { trayIcon.Visible = false; } catch {}
        try { SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS); } catch {}
        System.Threading.Timer exitDelay = null;
        exitDelay = new System.Threading.Timer(_ => {
            try { if (exitDelay != null) exitDelay.Dispose(); } catch { }
            try { Application.Exit(); } catch {}
        }, null, 2500, System.Threading.Timeout.Infinite);
    }
}
