using System;
using System.Runtime.InteropServices;

namespace Windows11Interop
{
    internal static class Guids
    {
        public static readonly Guid CLSID_ImmersiveShell =
             new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239" );
        public static readonly Guid CLSID_VirtualDesktopManagerInternal =
            new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B" );
        public static readonly Guid CLSID_VirtualDesktopManager =
            new Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A");
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid( "3F07F4BE-B107-441A-AF0F-39D82529072C" )]
    //[Guid("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")]
    internal interface IVirtualDesktop
    {
        void notimpl1(); // void IsViewVisible(IApplicationView view, out int visible);
        Guid GetId();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid( "53F5CA0B-158F-4124-900C-057158060B27" )]
    //[Guid("F31574D6-B682-4CDC-BD56-1827860ABEC6")]
    internal interface IVirtualDesktopManagerInternal
    {
        int GetCount();
        void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
        void notimpl2();  // void CanViewMoveDesktops(IApplicationView view, out int itcan);
        IVirtualDesktop GetCurrentDesktop();
        void GetDesktops(out IObjectArray desktops);
        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
        void SwitchDesktop(IVirtualDesktop desktop);
        IVirtualDesktop CreateDesktop();
        void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
        IVirtualDesktop FindDesktop(ref Guid desktopid);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("a5cd92ff-29be-454c-8d04-d82879fb3f1b")]
    internal interface IVirtualDesktopManager
    {
        int IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
        Guid GetWindowDesktopId(IntPtr topLevelWindow);
        void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4bba-A805-5E9F541BD8C9")]
    internal interface IObjectArray
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object obj);
    }


    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    internal interface IServiceProvider10
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object QueryService(ref Guid service, ref Guid riid);
    }


    [StructLayout( LayoutKind.Sequential )]
    internal struct Size
    {
        public int X;
        public int Y;
    }

    [StructLayout( LayoutKind.Sequential )]
    internal struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    internal enum APPLICATION_VIEW_CLOAK_TYPE : int
    {
        AVCT_NONE            = 0,
        AVCT_DEFAULT         = 1,
        AVCT_VIRTUAL_DESKTOP = 2
    }

    internal enum APPLICATION_VIEW_COMPATIBILITY_POLICY : int
    {
        AVCP_NONE                = 0,
        AVCP_SMALL_SCREEN        = 1,
        AVCP_TABLET_SMALL_SCREEN = 2,
        AVCP_VERY_SMALL_SCREEN   = 3,
        AVCP_HIGH_SCALE_FACTOR   = 4
    }

    [ComImport]
    [InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
    [Guid( "372E1D3B-38D3-42E4-A15B-8AB2B178F513" )]
    internal interface IApplicationView
    {
        void GetIIdsSlot();
        void GetRuntimeClassNameSlot();
        void GetTrustLevelSlot();
        int  SetFocus();
        int  SwitchTo();
        int  TryInvokeBack( IntPtr /* IAsyncCallback* */ callback );
        int  GetThumbnailWindow( out IntPtr hWnd );
        int  GetMonitor( out         IntPtr /* IImmersiveMonitor */ immersiveMonitor );
        int  GetVisibility( out      int visibility );
        int  SetCloak( APPLICATION_VIEW_CLOAK_TYPE cloakType, int unknown );
        int  GetPosition( ref Guid guid /* GUID for IApplicationViewPosition */, out IntPtr /* IApplicationViewPosition** */ position );
        int  SetPosition( ref IntPtr /* IApplicationViewPosition* */ position );
        int  InsertAfterWindow( IntPtr hWnd );
        int  GetExtendedFramePosition( out                              Rect rect );
        int  GetAppUserModelId( [MarshalAs( UnmanagedType.LPWStr )] out string id );
        int  SetAppUserModelId( string id );
        int  IsEqualByAppUserModelId( string id, out int result );
        int  GetViewState( out uint state );
        int  SetViewState( uint state );
        int  GetNeediness( out               int neediness );
        int  GetLastActivationTimestamp( out ulong timestamp );
        int  SetLastActivationTimestamp( ulong timestamp );
        int  GetVirtualDesktopId( out Guid guid );
        int  SetVirtualDesktopId( ref Guid guid );
        int  GetShowInSwitchers( out  int flag );
        int  SetShowInSwitchers( int flag );
        int  GetScaleFactor( out             int factor );
        int  CanReceiveInput( out            bool canReceiveInput );
        int  GetCompatibilityPolicyType( out APPLICATION_VIEW_COMPATIBILITY_POLICY flags );
        int  SetCompatibilityPolicyType( APPLICATION_VIEW_COMPATIBILITY_POLICY flags );
        int  GetSizeConstraints( IntPtr /* IImmersiveMonitor* */ monitor, out Size size1, out Size size2 );
        int  GetSizeConstraintsForDpi( uint uint1, out Size size1, out Size size2 );
        int  SetSizeConstraintsForDpi( ref uint uint1, ref Size size1, ref Size size2 );
        int  OnMinSizePreferencesUpdated( IntPtr hWnd );
        int  ApplyOperation( IntPtr /* IApplicationViewOperation* */ operation );
        int  IsTray( out                  bool isTray );
        int  IsInHighZOrderBand( out      bool isInHighZOrderBand );
        int  IsSplashScreenPresented( out bool isSplashScreenPresented );
        int  Flash();
        int  GetRootSwitchableOwner( out                              IApplicationView rootSwitchableOwner );
        int  EnumerateOwnershipTree( out                              IObjectArray     ownershipTree );
        int  GetEnterpriseId( [MarshalAs( UnmanagedType.LPWStr )] out string           enterpriseId );
        int  IsMirrored( out                                          bool             isMirrored );
        int  Unknown1( out                                            int              unknown );
        int  Unknown2( out                                            int              unknown );
        int  Unknown3( out                                            int              unknown );
        int  Unknown4( out                                            int              unknown );
        int  Unknown5( out                                            int              unknown );
        int  Unknown6( int                                                             unknown );
        int  Unknown7();
        int  Unknown8( out int   unknown );
        int  Unknown9( int       unknown );
        int  Unknown10( int      unknownX, int unknownY );
        int  Unknown11( int      unknown );
        int  Unknown12( out Size size1 );
    }


    [ComImport]
    [InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
    [Guid( "1841C6D7-4F9D-42C0-AF41-8747538F10E5" )]
    internal interface IApplicationViewCollection
    {
        int  GetViews( out         IObjectArray array );
        int  GetViewsByZOrder( out IObjectArray array );
        int  GetViewsByAppUserModelId( string   id,          out IObjectArray     array );
        int  GetViewForHWnd( IntPtr             hWnd,        out IApplicationView view );
        int  GetViewForApplication( object      application, out IApplicationView view );
        int  GetViewForAppUserModelId( string   id,          out IApplicationView view );
        int  GetViewInFocus( out IntPtr         view );
        int  Unknown1( out       IntPtr         view );
        void RefreshCollection();
        int  RegisterForApplicationViewChanges( object listener, out int cookie );
        int  UnregisterForApplicationViewChanges( int  cookie );
    }

}