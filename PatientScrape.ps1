function Get-Patient {
# https://stackoverflow.com/questions/25369285/how-can-i-get-all-window-handles-by-a-process-in-powershell
BEGIN {
    function Get-WindowText($hwnd) {
        $buffer_len = 256
        $sb = New-Object text.stringbuilder -ArgumentList ($buffer_len)
        $WM_GETTEXT = 0x000D
        $rtnlen = [apifuncs]::SendMessage($hwnd, $WM_GETTEXT, $buffer_len, $sb)
        $sb.tostring()
    }

    if ($null -eq ("APIFuncs" -as [type])) {
        Add-Type  @"
        using System;
        using System.Runtime.InteropServices;
        using System.Collections.Generic;
        using System.Text;
        using System.Text.RegularExpressions;
        public class APIFuncs {
            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
            public static extern int GetWindowText(IntPtr hwnd, StringBuilder lpString, int cch);

            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
            public static extern Int32 GetWindowTextLength(IntPtr hWnd);

            [DllImport("User32.dll")]
            public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, StringBuilder lParam);

            [DllImport("user32")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

            public static List<IntPtr> GetChildWindows(IntPtr parent) {
                List<IntPtr> result = new List<IntPtr>();
                GCHandle listHandle = GCHandle.Alloc(result);
                try {
                    EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                    EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
                } finally {
                    if (listHandle.IsAllocated)
                        listHandle.Free();
                }
               return result;
            }

            private static bool EnumWindow(IntPtr handle, IntPtr pointer) {
                GCHandle gch = GCHandle.FromIntPtr(pointer);
                List<IntPtr> list = gch.Target as List<IntPtr>;
                if (list == null) {
                    throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
                }
                list.Add(handle);
                return true;
            }
            public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

            public delegate bool CallBackPtr(IntPtr hwnd, IntPtr pointer);
            
            public static bool Report(IntPtr hwnd, IntPtr pointer) {
                GCHandle gch = GCHandle.FromIntPtr(pointer);
                List<IntPtr> results = gch.Target as List<IntPtr>;

                StringBuilder sb = new StringBuilder("", 256);
                int result = GetWindowText(hwnd, sb, sb.Capacity);
                
                String title = sb.ToString();
                Regex r = new Regex(@"^Edit patient");

                if (title.Length > 0 && r.IsMatch(title)) {
                    results.Add(hwnd);
                }
                return true;
            }

            [DllImport("user32.dll")]
            private static extern int EnumWindows(CallBackPtr callPtr, IntPtr pointer);

            public static List<IntPtr> main() {
                List<IntPtr> results = new List<IntPtr>();
                GCHandle listHandle = GCHandle.Alloc(results);
                IntPtr yo = GCHandle.ToIntPtr(listHandle);
                CallBackPtr callBackPtr = new CallBackPtr(Report);
                EnumWindows(callBackPtr, GCHandle.ToIntPtr(listHandle));
                return results;
            }
        }
"@
    }
}

    PROCESS {
        $sn = ([apifuncs]::main())
        if ($sn.Count -eq 1) {
            $editHwnd = $sn[0]
            $children = ([apifuncs]::GetChildWindows($editHwnd))
            $lastName = (Get-WindowText($children[1]))
            $firstName = (Get-WindowText($children[2]))
            $dob = (Get-WindowText($children[5]))
            $email = (Get-WindowText($children[25]))
            $split = $dob.Split("/")
            $day = $split[0]
            if ($day.Length -eq 1) {
                $day = "0" + $day
            }
            $parsedDOB = $day + $split[1] + $split[2].Substring(2,2)
            $generalNotes = (Get-WindowText($children[70]))
            $apptNotes = (Get-WindowText($children[71]))
            $noZip = $False
            if ($generalNotes -match ".*zip.*" -or $generalNotes -match ".*unprotected.*" -or $generalNotes -match ".*password.*") {
                $noZip = $true
            }
            elseif ($apptNotes -match ".*zip.*" -or $apptNotes -match ".*unprotected.*" -or $apptNotes -match ".*password.*") {
                $noZip = $true
            }
            return $email, $parsedDOB, $firstName, $lastName, $editHwnd, $noZip
        }
        return $sn.Count
    }
}
