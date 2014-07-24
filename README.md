portlist
========

Port List - lists available (or remembered) serial or printer ports on MS Windows

One problem with connecting a new USB or Bluetooth to serial or parallel device to a
PC running Microsoft Windows is that you don't always see the message flash up that
tells you the serial or parallel port name. You may forget the name, or the name may
change if you plug it into a different USB port, or the program using the port did
not close the port before you reconnect the device.

This also affects things that pretend to be a serial or parallel port over USB,
but microcontroller development devices such as the Arduino.

I found some programs that claimed to help with this, but they were flawed in not
finding all my different devices or took a long time by literally trying to open
(& then close) tens of ports named COM1 and incrementing the number.

Hence this Portlist program, which tries to be fast and comprehensive. It finds
the available parallel and/or serial ports through querying the registry and device
setup API. It is also able to find and list all such ports known to Windows, and
remembered for devices that are not currently available. Amongst other features,
the reported list can also be limited to selected USB VIDs (Vendor Ids) or VID:PID
pairs (Vendor & Product Ids).


As of the first public release, version 0.9, the available options reported through
'portlist /h' are:

    portlist - COM & LPT port listing utility - version 0.9
	        Copyright (c) 2013, 2014 Anthony Naggs

    Usage: portlist [-a] [-c] [-h|-?] [-l] [-pid <vid>:<pid>] [-vid <vid>] [-v] [-x] [-xc] [-xl]
            -a                list available (default) and also remembered ports
            -c                show Copyright and Warranty details (GNU General Public License v2)
			-h or -?          show this help text plus examples
			-l                long including Bus type, Vendor & Product IDs
			-vid <vid>        specify a hex Vendor ID to match
			-pid <vid>:<pid>  specify pair of hex Vendor & Product IDs to match
			-v                verbose multi-line per port list (implies -l)
			-x                exclude available ports => list only remembered ports
			-xc               exclude COM ports
			-xl               exclude LPT/PRN ports
    Note that multiple '-vid' and/or '-pid' parameters can be specified.
    Options can start with / or - and be upper or lowercase.

    Examples:
			portlist                    : list available ports and description
			portlist -l                 : longer, detailed list of available ports
			portlist -a                 : all available & remembered ports
			portlist /XL                : exclude printer ports => COM ports only
			portlist -pid 2341:0001     : match Arduino Uno VID/PID
			portlist /pid 04d8:000A     : match Microchip USB serial port ref
			portlist -pid 1d50:6098     : match Aperture Labs' RFIDler
			portlist /vid 0403          : match FTDI's Vendor ID (to find USB serial bridge chips)
			portlist -vid 4e8 -vid 421  : match either Samsung or Nokia VIDs

I hope you find this program useful, and welcome feedback on bugs & omissions.

As of portlist version 0.9 the program is compiled with Microsoft Visual Studio 2010.

GPL v2 Copyright
================

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

Examples of portlist usage
==========================

Examples show two Microchip based dev boards connected, a USB Bit Whacker (UBW32)
with the default Microchip USB to serial VID & PID, and an Aperture Labs
RFIDler LF prototype.

Normal brief mode
Note that A (Active) column is only shown when -a option is given to include
remebered as well as active ports in the list.

	C:\>portlist
	Port   Description
	COM3   USB Serial (UBW-based) communi
	COM4   USB Serial Port
    
	2 ports found.

	C:\>portlist -a
	Port   A Description
	COM3   Y USB Serial (UBW-based) communi
	COM4   Y USB Serial Port
    
	2 ports found.

Longer 1 line per port mode.
Note again that A (Active) column is only shown when -a option is given to
include remembered as well as active ports in the list.

	C:\>portlist -l
	Port    Bus    VID  PID  Rev  Product, Vendor
	COM3    USB    04D8 000A 0100 USB Serial (UBW-based) communi, SchmalzHaus LLC
	COM4    USB    1D50 6098 0020 USB Serial Port, Aperture Labs Ltd.

	2 ports found.

	C:\>portlist -l -a
	Port    A Bus    VID  PID  Rev  Product, Vendor
	COM3    Y USB    04D8 000A 0100 USB Serial (UBW-based) communi, SchmalzHaus LLC
	COM4    Y USB    1D50 6098 0020 USB Serial Port, Aperture Labs Ltd.

	2 ports found.

Verbose mode

	C:\>portlist -v
	Port    A Bus    VID  PID  Rev  Product, Vendor
	COM3    Y USB    04D8 000A 0100 USB Serial (UBW-based) communi, SchmalzHaus LLC
			  Device Class: Ports
			  Hardware Id: USB\VID_04D8&PID_000A&REV_0100
			  Physical Device Object: \Device\USBPDO-7
			  Location Info: Port_#0001.Hub_#0001

	COM4    Y USB    1D50 6098 0020 USB Serial Port, Aperture Labs Ltd.
			  Device Class: Ports
			  Hardware Id: USB\VID_1D50&PID_6098&REV_0020
			  Physical Device Object: \Device\USBPDO-8
			  Location Info: Port_#0001.Hub_#0004

    2 ports found.

