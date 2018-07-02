# portlist

Port List - lists available (optionally all remembered) COM (serial or modem)
or LPT (printer) ports on Microsoft Windows 2000 or later (XP, Vista, 7, 8 or 10).

## purpose

One problem with connecting a new USB or Bluetooth (etc..) serial or parallel
device to a PC running Microsoft Windows is that you do not always see the
message flash up that tells you the serial or parallel port name. You may
forget the name, or the name may change if you plug it into a different USB
port, or the program using the port did not close the port before you
reconnect the device.

This also affects things that pretend to be a serial or parallel port over
USB; GPS receivers, or microcontroller development devices such as an Arduino.

Amongst other features the reported list can also be limited to selected USB
VIDs (Vendor Ids) or VID:PID pairs (Vendor & Product Ids). Long (-l) and
verbose (-v) modes report additional information beyond a simple list of
ports, helpful for debugging driver .inf files or the microcontroller
development board at the other end of the USB connection.

Example output for each listing mode is towards the end of this readme.


As of version 0.9.1, the available options reported through 'portlist -h' are:

	portlist - COM & LPT port listing utility - version 0.9.1
			Copyright (c) 2013, 2014 Anthony Naggs

	portlist comes with ABSOLUTELY NO WARRANTY.
	This software is free, you are welcome to redistribute it under certain conditions.
	Type 'portlist -c' for Copyright, Warranty and distribution details.
	portlist source and binary files are available from:
			https://github.com/tonynaggs/portlist

	Usage:
	portlist [-a] [-l] [-usb[=<vid>[:<pid>]] [-v] [-x] [-xc] [-xl]
	portlist [-c] [-h|-?]
			-a                list all: available (default) plus remembered ports
			-c                show GPL Copyright and Warranty details
			-h or -?          show this help text plus examples
			-l                long including Bus type, Vendor & Product IDs
			-v                verbose multi-line per port list (implies -l)
			-usb              specify that any USB devices match
			-usb=<vid>        specify a USB Vendor ID (in hex) to match
			-usb=<vid>:<pid>  pair of USB Vendor & Product IDs (in hex) to match
			-x                exclude available ports => list only remembered ports
			-xc               exclude COM ports
			-xl               exclude LPT/PRN ports
			Notes: Multiple '-usb' parameters can be specified.
			Options can start with / or - and be upper or lowercase.

	Examples:
			portlist                    : list available ports and description
			portlist -l                 : longer, detailed list of available ports
			portlist -a                 : all available & remembered ports
			portlist /XL                : exclude printer ports => COM ports only
			portlist -usb=2341:0001     : match Arduino Uno VID/PID
			portlist /usb=04d8:000A     : match Microchip USB serial port ref
			portlist -usb=1d50:6098     : match Aperture Labs' RFIDler
			portlist /usb=0403          : match FTDI's Vendor ID (eg serial bridges)
			portlist -usb=4e8 -usb=421  : match either Samsung or Nokia VIDs


I hope you find this program useful, and welcome feedback on bugs & omissions.

As of portlist version 0.9 the program is compiled with Microsoft Visual Studio 2010,
and built for Intel x86 32-bit architecture.

## Bug reporting

If reporting bugs please indicate which Windows version (200, XP, Vista, 7, 8, or 10) you are using.
The program is not tested with earlier Windows version, and will probably fail.


# GPL v2 Copyright

Limited assignment of rights through the GNU General Public License version 2
is described below:

	This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.
    
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    
    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

# Examples of portlist usage

Examples show two Microchip based dev boards connected, a USB Bit Whacker (UBW32)
with the default Microchip USB to serial VID & PID, and an Aperture Labs
RFIDler LF prototype. In all these examples the overlong UBW32 is shortened
when it is displayed, in order to fit on one line.

**Normal**, brief mode.

	C:\>portlist
	Port   Friendly name
	COM3   USB Serial (UBW-based) communications port (COM3)
	COM4   USB Serial Port (COM4)

	2 ports found.

**All** ports (-a option), including names reserved for (currently) unavailable ports.
Note that this adds the A (Active).

	C:\>portlist -a
	Port   A Friendly name
	COM3   A USB Serial (UBW-based) communications port (COM3)
	COM4   A USB Serial Port (COM4)

	2 ports found.

**Longer** report (-l option), 1 line per port.
Can be combined with the -a option to include remembered as well as active
ports in the list, and add the A (active) column.

	C:\>portlist -l
	Port   Bus    VID  PID  Rev  Friendly name
	COM3   USB    04D8 000A 0100 USB Serial (UBW-based) communications port (COM3)
	COM4   USB    1D50 6098 0020 USB Serial Port (COM4)

	2 ports found.


**Verbose** mode, -v option, finds and reports details useful for debugging whether
the device is correctly recognised by windows and the correct drivers are installed.

Note that 'Physical Device Object' only exists for ports that are currently
available.

	C:\>portlist -v
	Port   A Bus    VID  PID  Rev  Friendly name
	COM3   A USB    04D8 000A 0100 USB Serial (UBW-based) communications port (COM3)
			 Vendor: SchmalzHaus LLC
			 Product: USB Serial (UBW-based) communications port
			 Device Class: Ports
			 Hardware Id: USB\VID_04D8&PID_000A&REV_0100
			 Physical Device Object: \Device\USBPDO-6
			 Location Info: Port_#0001.Hub_#0001

	COM4   A USB    1D50 6098 0020 USB Serial Port (COM4)
			 Vendor: Aperture Labs Ltd.
			 Product: USB Serial Port
			 Device Class: Ports
			 Hardware Id: USB\VID_1D50&PID_6098&REV_0020
			 Physical Device Object: \Device\USBPDO-7
			 Location Info: Port_#0001.Hub_#0004

	2 ports found.

# Rationale for writing yet another port listing program

I found some programs that claimed to do this, but they were flawed in some way,
such as only finding some of my different devices e.g. my modem or Bluetooth connected
devices were often missed.
Alternatively I had wait minutes as they literally tried to open (& then close) tens of
ports named COM1, COM2, COM3 etc ... There are lots of reasons this can take a long time,
a laptop remembers ports that are in its dock, or a Bluetooth device may be out of rane.

Hence this *portlist* program, which tries to be fast and comprehensive. It finds the 
available parallel and/or serial ports by querying the Windows registry and using the
device setup API. It is also able to find and list all (option -a) ports
that are known to Windows and remembered for devices that are not currently available.

