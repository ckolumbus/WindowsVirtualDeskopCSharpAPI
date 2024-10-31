// source: https://stackoverflow.com/questions/32416843/programmatic-control-of-virtual-desktops-in-windows-10

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Flow.Launcher.Plugin.OpenWindowSearch;
using Windows11Interop;

namespace WindowsInterop
{
    public class Desktop
    {
        public static int Count
        {
            // Returns the number of desktops
            get { return DesktopManager.Manager.GetCount(); }
        }

        public static Desktop Current
        {
            // Returns current desktop
            get { return new Desktop(DesktopManager.Manager.GetCurrentDesktop()); }
        }

        public static Desktop FromIndex(int index)
        {
            // Create desktop object from index 0..Count-1
            return new Desktop(DesktopManager.GetDesktop(index));
        }

        public static List<Guid> Desktops(){
            return DesktopManager.GetDesktops();
        }

        public static Guid GuidFromWindow(IntPtr hWnd)
        {
            // return desktop object to desktop on which window <hWnd> is displayed
            if ( hWnd == IntPtr.Zero ) throw new ArgumentNullException();

            // Creates desktop object on which window <hWnd> is displayed
            return DesktopManager.WManager.GetWindowDesktopId(hWnd);
        }
        public static Desktop FromWindow(IntPtr hWnd)
        {
            // return desktop object to desktop on which window <hWnd> is displayed
            if ( hWnd == IntPtr.Zero ) throw new ArgumentNullException();

            // Creates desktop object on which window <hWnd> is displayed
            var id = DesktopManager.WManager.GetWindowDesktopId(hWnd);

            // TODO: FindDesktop throws an AccessViolationException on Windows 11, don't know why
            //       temp workaround: use GuidFromWindow instead to only get the Guid
            return new Desktop(DesktopManager.Manager.FindDesktop(ref id));
        }

        public static Desktop Create()
        {
            // Create a new desktop
            return new Desktop(DesktopManager.Manager.CreateDesktop());
        }

        public void Remove(Desktop fallback = null)
        {
            // Destroy desktop and switch to <fallback>
            var back = fallback == null ? DesktopManager.GetDesktop(0) : fallback.itf;
            DesktopManager.Manager.RemoveDesktop(itf, back);
        }

        public bool IsVisible
        {
            // Returns <true> if this desktop is the current displayed one
            get { return object.ReferenceEquals(itf, DesktopManager.Manager.GetCurrentDesktop()); }
        }

        public void MakeVisible()
        {
            // Make this desktop visible
            DesktopManager.Manager.SwitchDesktop(itf);
        }

        public Desktop Left
        {
            // Returns desktop at the left of this one, null if none
            get
            {
                IVirtualDesktop desktop;
                int hr = DesktopManager.Manager.GetAdjacentDesktop(itf, 3, out desktop);
                if (hr == 0) return new Desktop(desktop);
                else return null;

            }
        }

        public Desktop Right
        {
            // Returns desktop at the right of this one, null if none
            get
            {
                IVirtualDesktop desktop;
                int hr = DesktopManager.Manager.GetAdjacentDesktop(itf, 4, out desktop);
                if (hr == 0) return new Desktop(desktop);
                else return null;
            }
        }

/*
        public void MoveWindow(IntPtr handle)
        {
            // Move window <handle> to this desktop
            DesktopManager.WManager.MoveWindowToDesktop(handle, itf.GetId());

            // doesn't work for out-or-process windows (== all).. need additional impl like
            // https://github.com/tmyt/VDMHelper/blob/eb5df82aa65ea85de2eca394bef637f8d7cbd848/VDMHelperCLR/src/VDMHelperCLR.cpp#L66
        }
*/
        public void MoveWindow( IntPtr hWnd )
        {
            // move window to this desktop
            if ( hWnd == IntPtr.Zero ) throw new ArgumentNullException();
            _ = NativeMethods.GetWindowThreadProcessId( hWnd, out var processId );

            if ( Process.GetCurrentProcess().Id == processId )
            {
                // window of process
                try // the easy way (if we are owner)
                {
                    DesktopManager.WManager.MoveWindowToDesktop( hWnd, itf.GetId() );
                }
                catch // window of process, but we are not the owner
                {
                    DesktopManager.AppViewCollection.GetViewForHWnd( hWnd, out var view );
                    DesktopManager.Manager.MoveViewToDesktop( view, itf);
                }
            }
            else
            {
                // window of other process
                DesktopManager.AppViewCollection.GetViewForHWnd( hWnd, out var view );
                try
                {
                    DesktopManager.Manager.MoveViewToDesktop( view, itf );
                }
                catch
                {
                    // could not move active window, try main window (or whatever windows thinks is the main window)
                    DesktopManager.AppViewCollection.GetViewForHWnd(
                        Process.GetProcessById( processId ).MainWindowHandle,
                        out view );
                    DesktopManager.Manager.MoveViewToDesktop( view, itf );
                }
            }
        }
        public Guid GetId()
        {
            return itf.GetId();
        }

        public bool HasWindow(IntPtr handle)
        {
            // Returns true if window <handle> is on this desktop
            return itf.GetId() == DesktopManager.WManager.GetWindowDesktopId(handle);
        }

        public override int GetHashCode()
        {
            return itf.GetHashCode();
        }
        public override bool Equals(object obj)
        {
            var desk = obj as Desktop;
            return desk != null && object.ReferenceEquals(this.itf, desk.itf);
        }

        private IVirtualDesktop itf;
        private Desktop(IVirtualDesktop itf) { this.itf = itf; }

        public Desktop()
        {
        }
    }

    internal static class DesktopManager
    {
        static DesktopManager()
        {
            var shell = (IServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_ImmersiveShell));

            Manager = (IVirtualDesktopManagerInternal)shell.QueryService(
                Guids.CLSID_VirtualDesktopManagerInternal,
                typeof( IVirtualDesktopManagerInternal ).GUID );

            WManager = (IVirtualDesktopManager)Activator.CreateInstance(
                Type.GetTypeFromCLSID( Guids.CLSID_VirtualDesktopManager ) );

            AppViewCollection = (IApplicationViewCollection)shell.QueryService(
                typeof( IApplicationViewCollection ).GUID,
                typeof( IApplicationViewCollection ).GUID );
        }

        internal static List<Guid> GetDesktops()
        {
            List<Guid> virtualDesktops = new List<Guid>();

            int count = Manager.GetCount();
            IObjectArray desktops;
            Manager.GetDesktops(out desktops);
            for (int i = 0; i < count; i++)
            {
                object objdesk;
                desktops.GetAt(i, typeof( IVirtualDesktop ).GUID, out objdesk);
                virtualDesktops.Add(((IVirtualDesktop)objdesk).GetId());
                Marshal.ReleaseComObject(objdesk);
            }
            return virtualDesktops;
        }

        internal static IVirtualDesktop GetDesktop(int index)
        {
            int count = Manager.GetCount();
            if (index < 0 || index >= count) throw new ArgumentOutOfRangeException("index");
            IObjectArray desktops;
            Manager.GetDesktops(out desktops);
            object objdesk;
            desktops.GetAt(index, typeof( IVirtualDesktop ).GUID, out objdesk);
            Marshal.ReleaseComObject(desktops);
            return (IVirtualDesktop)objdesk;
        }

        internal static IVirtualDesktopManagerInternal Manager;
        internal static IVirtualDesktopManager WManager;
        internal static IApplicationViewCollection AppViewCollection;
    }
}