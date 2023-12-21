# RadioMan

![image](https://github.com/fafalone/RadioMan/assets/7834493/c0dc270a-2bcd-4002-bb7c-28c18441f739)

## Windows Radio Management

This app exposes undocumented radio management functionality in Windows. I've always been a big fan of making your own settings apps, so don't like being told the official one is the only way to do something. The `IMediaRadioManager`, `IRadioInstance`, and `IRadioCollection` interfaces are documented, but the coclasses representing the system objects actually implementing them are not; they're just documented for hardware providers to provide. There's three of these I could find: the WWAN manager, which includes cellular radios, the WLAN manager, which includes WiFi radios, and the Bluetooth manager (self explanatory).  These provide control over individual radios. I tested this app on my Surface tablet, where I had one of each. The app lets you control them all individually; you choose one on the list then can query, enable, or disable:

![image](https://github.com/fafalone/RadioMan/assets/7834493/00ab882e-38c4-424f-90d3-40f502a16a9d)

In the above picture, I first selected the radio then click Query to confirm it was on, then clicked disable. You can see in the system tray the WiFi has been disabled.

In addition, there's the `IRadioManager` interface and `RadioManagementAPI` coclass-- the master switch for all system radios, popularly known as Airplane Mode. Both of these are undocumented, and have only been partially reverse engineered. I found an example of using this interface in the [wintouchg repo](https://github.com/GrieferAtWork/wintouchg) here on GitHub. You can query the state and enable and disable. Below, you can see that after clicking Enable, the Airplane Mode indicator replaced the WiFi connection in the system tray, indicating we've successfully entered the mode:

![image](https://github.com/fafalone/RadioMan/assets/7834493/8a1978b5-88d3-4d19-ae64-d6063f169639)


## Requirements

-The `IMediaRadioManager` individual interfaces require Windows 8 or later. I'm unsure of when support for the `IRadioManager` airplane mode interface started, but can say it's on Windows 10 and later.

This project was written exclusively in [twinBASIC](https://twinbasic.com/) ([GitHub](https://github.com/twinbasic/twinbasic)), the currently in-development succesor to VB6 with full backwards compatibility. Written to be compiled for both 32bit and 64bit targets.

-All of the interfaces and APIs are part of my WinDevLib (Windows Development Library, formerly tbShellLib) package for twinBASIC; this is entirely downloaded into the project file (for now), which is why the size is so large. But the compiler takes only what it needs; the exe doesn't contain the whole thing. You can add it to your own projects from the Package Manager or from [the project repository](https://github.com/fafalone/WinDevLib).

-Any recent version of tB should compile it.

## How it works

The individual radios are a pretty straightforward object which provides a collection, which provides an individual interface:

```vba
[ InterfaceId ("70AA1C9E-F2B4-4C61-86D3-6B9FB75FD1A2") ]
[ OleAutomation (False) ]
Interface IRadioInstance Extends IUnknown
    Sub GetRadioManagerSignature(pguidSignature As UUID)
    Sub GetInstanceSignature(pbstrId As String)
    Sub GetFriendlyName(ByVal lcid As Long, pbstrName As String)
    Sub GetRadioState(pRadioState As DEVICE_RADIO_STATE)
    Sub SetRadioState(ByVal radioState As DEVICE_RADIO_STATE, ByVal uTimeoutSec As Long)
    [ PreserveSig ] Function IsMultiComm() As BOOL
    [ PreserveSig ] Function IsAssociatingDevice() As BOOL
End Interface
```

We just enumerate those:

```vba
    Set pWW = New WwanRadioManager
    If pWW IsNot Nothing Then
        pWW.GetRadioInstances(pColWW)
        If (pColWW IsNot Nothing) Then
            pColWW.GetCount(nWW)
            If nWW Then
                For i = 0 To nWW - 1
                    pColWW.GetAt(i, pInstWW)
                    Debug.Print "Add WW radio 0"
                    pInstWW.GetFriendlyName(GetUserDefaultLCID(), sBuff)
                    pInstWW.GetInstanceSignature(sBuffId)
                    pInstWW.GetRadioManagerSignature(priid)
                    List1.AddItem(sBuff & " (InstId=" & sBuffId & ", RmSig=" & GUIDToString(priid))
                    Set pInstWW = Nothing
                Next
            End If
            Text1.Text = CStr(nWW) & " WWAN " & IIf(nWW = 1, "radio", "radios") & " found."
```

and repeat for the other two. Once we have the `IRadioInstance`, `GetRadioState` and `SetRadioState` are very simple methods to use.



`IRadioManager` for airplane mode is even simpler:

```vba
[ InterfaceId ("db3afbfb-08e6-46c6-aa70-bf9a34c30ab7") ]
Interface IRadioManager Extends IUnknown
    Sub IsRMSupported(pdwState As Long) 'Always 1
    Sub GetUIRadioInstances(ppInstances As IUnknown) 'IUIRadioInstanceCollection
    Sub GetSystemRadioState(pbEnabled As BOOL, param2 As Long, param3 As RADIO_CHANGE_REASON)
    Sub SetSystemRadioState(ByVal bEnabled As BOOL)
    Sub Refresh()
    Sub OnHardwareSliderChange(ByVal param1 As Long, ByVal param2 As Long)
End Interface
```

The arguments besides enabled aren't well understood; but they don't seem critical to functionality. All we have to do is `Set pRM = New RadioManagementAPI` then we have the `IRadioManager` interface and the dead simple `GetSystemRadioState` and `SetSystemRadioState`.

Overall, this is a very simple program. The hard part was done by the people who reverse engineered this stuff to make it easy to consume by people like me. Simple, but very useful if like me you want alternative to the crappy Settings app Microsoft refuses to improve. Enjoy!

