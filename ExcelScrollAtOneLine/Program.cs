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

        public static void Main(string[] args)
        {
            const int spiGetwheelscrolllines = 0x0068;
            const int spiSetwheelscrolllines = 0x0069;

            uint defaultWheelSpeed;
            SystemParametersInfoA(spiGetwheelscrolllines, 0, out defaultWheelSpeed, 0);
            uint curWheelSpeed = defaultWheelSpeed;

            StringBuilder sb = new StringBuilder(1024);

            while (true)
            {
                GetWindowText(GetForegroundWindow(), sb, sb.Capacity);

                switch (sb.ToString().Contains(" - Excel"))
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