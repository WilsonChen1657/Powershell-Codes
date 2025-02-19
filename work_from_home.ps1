# Define the necessary Win32 API functions
Add-Type @"
using System;
using System.Runtime.InteropServices;

public struct MPOINT
{
    public int X;
    public int Y;
}

public class MouseTracker
{
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(ref MPOINT lpPoint);
    
    public static MPOINT GetMousePosition()
    {
        MPOINT point = new MPOINT();
        GetCursorPos(ref point);
        return point;
    }
}

public class MouseMover
{
    [DllImport("user32.dll")]
    public static extern void mouse_event(int dwFlags, uint dx, uint dy, uint dwData, IntPtr dwExtraInfo);

    public const int MOUSEEVENTF_MOVE = 0x0001;
    
    public static void MoveMouse(int x, int y)
    {
        mouse_event(MOUSEEVENTF_MOVE, (uint)x, (uint)y, 0, IntPtr.Zero);
    }
}
"@

<#
# Move mouse to the x=500, y=300 screen position
[MouseMover]::MoveMouse(500, 300)

# Get the mouse position
$mousePosition = [MouseTracker]::GetMousePosition()
#>

# Set the interval for moving the mouse (15 minutes in milliseconds)
$interval = 14 * 60 * 1000

# Loop to move the mouse every 15 minutes
while ($true) {
    # Get the mouse position
    $mousePosition = [MouseTracker]::GetMousePosition()
    if ( $mousePosition.X -eq 1919 -or $mousePosition.Y -eq 1079) {
        [MouseMover]::MoveMouse(-1920, -1080)
    }
    else {
        # Move the mouse slightly (you can adjust the values)
        [MouseMover]::MoveMouse(1, 1)  # Move mouse 1 pixel right and 1 pixel down
    }
    
    Start-Sleep -Milliseconds $interval  # Wait for 15 minutes
}
