using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelScrollAtOneLine
{
    internal class Program
    {
        [DllImport("User32.dll")]
        public static extern bool SystemParametersInfoA(uint uiAction, uint uiParam, out uint pvParam, uint fWinIni);

        [DllImport("User32.dll")]
        public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, uint pvParam, uint fWinIni);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        //[DllImport("user32.dll")]
        //private static extern int GetWindowThreadProcessId(IntPtr handle, out uint processId);

        //private static string GetProcessPath(IntPtr hwnd)
        //{
        //    uint pid;
        //    GetWindowThreadProcessId(hwnd, out pid);

        //    if (pid != 0)
        //    {
        //        Process proc = Process.GetProcessById((int)pid);
        //                        Console.WriteLine(proc);
        //        return proc.MainModule.FileName.ToString().ToLower();
        //    }

        //    return "";
        //}

        public static void Main(string[] args)
        {
            const int spiGetwheelscrolllines = 0x0068;
            const int spiSetwheelscrolllines = 0x0069;

            uint defaultWheelSpeed;
            SystemParametersInfoA(spiGetwheelscrolllines, 0, out defaultWheelSpeed, 0);
            uint curWheelSpeed = defaultWheelSpeed;

            //string processPath;

            StringBuilder stringBuilder = new StringBuilder(1024);
            string windowText;

            while (true)
            {
                //processPath = GetProcessPath(GetForegroundWindow());
                GetWindowText(GetForegroundWindow(), stringBuilder, stringBuilder.Capacity);
                windowText = stringBuilder.ToString();

                //switch (processPath.Length > 9 && processPath.Substring(processPath.Length - 9) == "excel.exe")
                switch (windowText.Length > 8 && windowText.Substring(windowText.Length - 8) == " - Excel")
                {
                    case true when curWheelSpeed != 1:
                        SystemParametersInfo(spiSetwheelscrolllines, 1, 0, 3);
                        curWheelSpeed = 1;
                        break;

                    case false when curWheelSpeed == 1:
                        SystemParametersInfo(spiSetwheelscrolllines, defaultWheelSpeed, 0, 3);
                        curWheelSpeed = defaultWheelSpeed;
                        break;
                }

                Thread.Sleep(20);
            }
        }
    }
}