/*
    portlist.c - list COM (serial) and/or LPT (parallel) ports on MS Windows 

    https://github.com/tonynaggs/portlist

    Copyright (c) 2013, 2014 Anthony Naggs. (tonynaggs@gmail.com) All rights reserved.

    Under the Copyright, Designs and Patents Act 1988 Anthony Naggs asserts his moral
    right to be identified as the author of this work.

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
*/


/*
    Coding style notes:
    Use of indentation and braces is similar to "K&R style", but with 4 spaces per level.
    All if / else statements should have {} around the conditional code.
    Indentation is always spaces not tabs, for readable code in different tools.

    History:
    2014-07-27 AMN v0.9.1 Minor fixes and code tidying.
                    Active column should only be shown when All (-a) or
                    Verbose (-v) flags are given by the user.
                    Print Windows Friendly Name for ports, only print
                    Vendor & product strings in (-v) verbose mode.
                    Improve handling & printing of port's bus name.
                    Include github location in source & copyright help text.
                    Add /usb commandline option.
                    Deprecate /vid & /pid commandline options.

    2014-07-16 AMN v0.9 Available for public beta.
                    Cleaned up code.
                    Rename to portlist.
                    Copyright changed to GPL v2.

    2013-11-22 AMN v0.2 First version working & shared.
                    Compiled with MSVC 2010.
                    Testing so far is limited to PCs I own.
 */


/* MS VC whines about some ANSI code & POSIX functions even when used safely
 * These defines suppress the messages and are needed before #include of header files
 */
#define _CRT_NONSTDC_NO_DEPRECATE
#define _CRT_SECURE_NO_WARNINGS

/*
 * This program needs to be linked with Setupapi.lib
 */
#pragma comment(lib,"Setupapi.lib")


// ANSI C headers
#include <stdio.h>
#include <stdlib.h>
#include <ctype.h>

// MS Windows platform headers, compatible with Windows 2000 and through Windows 8.1
#include <windows.h>
#include <devguid.h>
#include <SetupAPI.h>



// #define these to configure Debug prints etc
#ifdef _DEBUG
// #define OPTIONS_DEBUG 
// #define DEBUG_DEV_PROPERTIES
#endif

// configure development or deprecated code
// #define SUPPORT_DEPRECATED_VID_OR_PID_OPTIONS
// #define SUPPORT_SERIAL_NUMBER_REPORTING


typedef unsigned Bool;
const unsigned False = 0;
const unsigned True = 1;

// common substrings collected for ease of maintenance
const wchar_t* progname_msg = L"portlist";
const wchar_t* version_msg = L"0.9.1";
const wchar_t* copyright_msg = L"Copyright (c) 2013, 2014 Anthony Naggs";
const wchar_t* homeurl_msg = L"https://github.com/tonynaggs/portlist";

// English messages for help & bad options
const wchar_t* long_copyright_msg =
    L"Limited assignment of rights through the GNU General Public License\n"
    L"version 2 is described below:\n"
    L"\n"
    L"This program is free software; you can redistribute it and/or modify\n"
    L"it under the terms of the GNU General Public License as published by\n"
    L"the Free Software Foundation; either version 2 of the License, or\n"
    L"(at your option) any later version.\n"
    L"\n"
    L"This program is distributed in the hope that it will be useful,\n"
    L"but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
    L"MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the\n"
    L"GNU General Public License for more details.\n"
    L"\n"
    L"You should have received a copy of the GNU General Public License along\n"
    L"with this program; if not, write to the Free Software Foundation, Inc.,\n"
    L"51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.\n\n";
    
const wchar_t* usage_msgs[] = 
{
    L"[-a] [-l] [-usb[=<vid>[:<pid>]] [-v] [-x] [-xc] [-xl]",
    L"[-c] [-h|-?]",
    NULL
};

const wchar_t* option_msgs[] = 
{
    L"-a                list all: available (default) plus remembered ports",
    L"-c                show GPL Copyright and Warranty details",
    L"-h or -?          show this help text plus examples",
    L"-l                long including Bus type, Vendor & Product IDs",
#if defined(SUPPORT_DEPRECATED_VID_OR_PID_OPTIONS)
    L"-vid <vid>        [deprecated] specify a USB Vendor ID (in hex) to match",
    L"-pid <vid>:<pid>  [deprecated] pair of USB Vendor & Product IDs (in hex) to match",
#endif
    L"-v                verbose multi-line per port list (implies -l)",
    L"-usb              specify that any USB devices match",
    L"-usb=<vid>        specify a USB Vendor ID (in hex) to match",
    L"-usb=<vid>:<pid>  pair of USB Vendor & Product IDs (in hex) to match",
    L"-x                exclude available ports => list only remembered ports",
    L"-xc               exclude COM ports",
    L"-xl               exclude LPT/PRN ports",
    L"Notes: Multiple '-usb' parameters can be specified.",
    L"Options can start with / or - and be upper or lowercase.",
    NULL
};

const wchar_t* example_msgs[] =
{
    L"                    : list available ports and description",
    L" -l                 : longer, detailed list of available ports",
    L" -a                 : all available & remembered ports",
    L" /XL                : exclude printer ports => COM ports only",
    L" -usb               : match any USB device",
    L" -usb=2341:0001     : match Arduino Uno VID/PID",
    L" /usb=04d8:000A     : match Microchip USB serial port ref",
    L" -usb=1d50:6098     : match Aperture Labs' RFIDler",
    L" /usb=0403          : match FTDI's Vendor ID (eg serial bridges)",
    L" -usb=4e8 -usb=421  : match either Samsung or Nokia VIDs",
#if defined(SUPPORT_DEPRECATED_VID_OR_PID_OPTIONS)
    L" -pid 2341:0001     : [deprecated] match Arduino Uno VID/PID",
    L" /pid 04d8:000A     : [deprecated] match Microchip USB serial port ref",
    L" -pid 1d50:6098     : [deprecated] match Aperture Labs' RFIDler",
    L" /vid 0403          : [deprecated] match FTDI's Vendor ID (maker of USB to serial bridges)",
    L" -vid 4e8 -vid 421  : [deprecated] match either Samsung or Nokia VIDs",
#endif
    NULL
};

// bit flags for option switches
#define OPT_FLAG_ALL                0x00000001
#define OPT_FLAG_LONGFORM           0x00000002
#define OPT_FLAG_VERBOSE            0x00000004
#define OPT_FLAG_USBMATCH_VID       0x00000010
#define OPT_FLAG_USBMATCH_PIDVID    0x00000020
#define OPT_FLAG_USBMATCH_ANY       0x00000040
#define OPT_FLAG_EXCLUDE_COM        0x00001000
#define OPT_FLAG_EXCLUDE_LPT        0x00002000
#define OPT_FLAG_EXCLUDE_AVAILABLE  0x00004000

#define OPT_FLAG_HELP_COPYRIGHT     0x40000000
#define OPT_FLAG_HELP               0x80000000

// if any of the USB matching options are specified
#define OPT_FLAG_USBMATCH_SPECIFIED (OPT_FLAG_USBMATCH_PIDVID | OPT_FLAG_USBMATCH_VID | OPT_FLAG_USBMATCH_ANY)

/*
    Notes on Vendor & Product Ids
    =============================
   
    Currently portlist only supports matching of USB Vendor & Product Ids.

    USB specifies 16 bit Vendor, Product and optionally Revision Ids.

    Other buses use different Id sizes, and Windows stores these in different
    formats so without test systems it is hard to confirm that matching would
    work.

    EISA Vendor Id is 3 ASCII characters coded A = 00001b -> Z = 11010b.
    Product Id is 8 bits, Revision is 8 bits.

    Cardbus (32 bit bus PC Card) through PCI bridge is the same as PCI.

    PCMCIA (16 bit bus PC Card)
        CISTPL_VERS_1 (Card Information Structure tuple version 1) has:
        Major version 8 bits
        Minor version 8 bits
        Manufacturer Name string ASCIIZ
        Product Name string ASCIIZ
        Additional Info string ASCIIZ

    PCI has 16 bit Vendor Id, Device Id, and 8 bit Revision Id.
    Vendor Id 0xFFFF => slot empty, Windows also treats 0x0000 as slot empty.
    Class Code Register is 24 bits, should be non-zero.

    Firewire defines a 24 bit Vendor Id.
*/

// bit flags for retrieved data
#define RETRIEVED_VENDOR_ID         0x00000001
#define RETRIEVED_PRODUCT_ID        0x00000002
#define RETRIEVED_REV               0x00000004
#define RETRIEVED_PORTADDRESS       0x00000008
#define RETRIEVED_INTERRUPT         0x00000010
#define RETRIEVED_PORTINDEX         0x00000020
#define RETRIEVED_INDEXED           0x00000040


////////////////////////////////////////////////
// struct definintions
////////////////////////////////////////////////

struct u32_list {
    unsigned*   ulist;
    unsigned    count;
    unsigned    max;
};


typedef struct portinfo {
    wchar_t*            portname;       // COM1, PRN, ...
    wchar_t*            friendlyname;   // Windows friendly name

    // sorting key info
    size_t              prefixlen;      // length of "COM", "LPT" prefix, or strlen of name
    unsigned            portnumber;     // upto 3 digit number following COM or LPT

    // optional info for long listing
    wchar_t*            busname;
    Bool                isUSB;          // probably USB, therefore VID matching is valid
    unsigned long       vendorId;
    unsigned long       productId;
    unsigned long       revision;
    unsigned            retrieved;      // bit flags

    // optional info for verbose listing
    Bool                isAvailable;
    wchar_t*            product;        // product description eg "USB Serial Port"
    wchar_t*            vendor;
    wchar_t*            hardwareid;
    wchar_t*            location;
    wchar_t*            physdevobj;
    wchar_t*            devclass;
#if defined(SUPPORT_SERIAL_NUMBER_REPORTING)
    wchar_t*            serialnumber;
#endif

    // verbose details from registry, for legacy ports (no Plug & Play)
    unsigned long       portaddress;
    unsigned long       interrupt;

    // verbose details from registry, for multi-port devices
    //unsigned long      multiport;   // MultiportDevice (REG_DWORD)
    unsigned long       portindex;   // PortIndex (REG_DWORD)
    unsigned long       indexed;     // Indexed (REG_DWORD), bool true if PortIndex is index rather than bitmap

    // for linked list
    struct portinfo*     next;
} PortInfo;


typedef struct portlist {
    unsigned        opt_flags;

    struct u32_list pidvidlist;  // list of USB VID:PID pairs (NB USB only as only 16 bits each)
    struct u32_list vidlist;     // list of naked USB VIDs

    PortInfo*       ports;       // linked list of brief port info
} PortList;


////////////////////////////////////////////////
// function prototypes
////////////////////////////////////////////////

void listadd(struct u32_list* list, unsigned value);
int errorprint(const wchar_t* message);
int errorprintf(const wchar_t* format, ...);
void usage(Bool help_examples, Bool help_copyright);
int matchoption(PortList* list, int argc, wchar_t** argv);
Bool checkoptions(PortList* list, int argc, wchar_t** argv);
Bool findinlist(struct u32_list list, unsigned value);
Bool checkpidandvidlists(PortList* list, PortInfo* pInfo);
void trygetdevice_regdword(HKEY devkey, wchar_t* keyname, DWORD* result, unsigned int* flags, unsigned int attribflag);
wchar_t* getportname(HKEY devkey);
void getverboseportinfo(HKEY devkey, PortInfo* pInfo);
PortInfo* getdevicesetupinfo(HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, unsigned opt_flags);
wchar_t* portstringproperty(HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, DWORD devprop);
wchar_t* wcs_dupsubstr(const wchar_t* string, size_t length);
Bool wcs_seekul(wchar_t** pString, const wchar_t* SubStr, unsigned long* pOutValue, int Radix);
Bool getportdetails(PortList* list, HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, PortInfo* pInfo);
int portcmp(PortInfo* p1, PortInfo* p2);
Bool getdeviceinfo(PortList* list, HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData);
unsigned listdevices(PortList* list, HDEVINFO hDevInfo);
unsigned listclass(PortList* list, CONST GUID *guid);
void listports(PortList* list);



void listadd(struct u32_list* list, unsigned value)
{
    if ((list->count + 1) > list->max) {
        size_t sz = list->max + 10 * sizeof(unsigned); // granularity
        list->ulist = (unsigned *) realloc(list->ulist, sz);
        if (list->ulist == NULL) {
            errorprint(L"Failed to parse input parameters - memory allocation failed!");
            exit(-1);
        }
    }

    list->ulist[list->count++] = value;
}


// print program name & error message to stderr
int errorprint(const wchar_t* message)
{
    return fwprintf(stderr, L"%s: %s\n", progname_msg, message);
}


// print program name & formatted error message to stderr
int errorprintf(const wchar_t* format, ...)
{
    int res;
    va_list arglist;

    va_start(arglist, format);

    res = fwprintf(stderr, L"%s: ", progname_msg);

    if (res > 0) {
        res += vfwprintf(stderr, format, arglist);
        fputws(L"\n", stderr);
        res += 1;
    }
    return res;
}


void usage(Bool help_examples, Bool help_copyright)
{
    fwprintf(stderr, L"%s - COM & LPT port listing utility - version %s\n\t%s\n\n",
        progname_msg, version_msg, copyright_msg);

    if (help_copyright) {
        fwprintf(stderr, L"\tHome URL %s\n\n", homeurl_msg);
        fputws(long_copyright_msg, stderr);
    } else {
        fwprintf(stderr, L"%s comes with ABSOLUTELY NO WARRANTY.\n"
            L"This software is free, you are welcome to redistribute it under certain conditions.\n"
            L"Type `%s -c' for Copyright, Warranty and distribution details.\n"
            L"%s source and binary files are available from:\n\t%s\n\n",
                progname_msg, progname_msg, progname_msg, homeurl_msg);
    }

    if (help_examples || !help_copyright) {
        const wchar_t** msgs;

        fputws(L"Usage:\n", stderr);

        // lines mentioning program name
        for (msgs = usage_msgs; *msgs; msgs++) {
            fwprintf(stderr, L"%s %s\n", progname_msg, *msgs);
        }

        // list of options
        for (msgs = option_msgs; *msgs; msgs++) {
            fwprintf(stderr, L"\t%s\n", *msgs);
        }

        fputws(L"\n", stderr);
    }

    if (help_examples) {
        const wchar_t** msgs = example_msgs;

        fputws(L"Examples:\n", stderr);
        for (; *msgs; msgs++) {
            fwprintf(stderr, L"\t%s%s\n", progname_msg, *msgs);
        }

        fputws(L"\n", stderr);
    }

    fflush(stderr);
}


struct opt_info {
    const wchar_t* opt_text;
    unsigned       set_flags;
    unsigned       clear_flags;
};

struct opt_info opt_list[] = {
    // -a                all known ports including those not currently available
    { L"a", OPT_FLAG_ALL, 0 },
    // -c                show GPL copyright
    { L"c", OPT_FLAG_HELP_COPYRIGHT, 0 },
    // -h or -?          show help text plus examples
    { L"h", OPT_FLAG_HELP, 0 },
    { L"?", OPT_FLAG_HELP, 0 },
    // -l                long including Bus type, Vendor & Product IDs
    { L"l", OPT_FLAG_LONGFORM, 0 },
    // -v                verbose output, Vendor string, ...
    { L"v", OPT_FLAG_VERBOSE | OPT_FLAG_LONGFORM, 0 },
    // -x                exclude available ports (list only remembered ports)
    { L"x", OPT_FLAG_EXCLUDE_AVAILABLE | OPT_FLAG_ALL, 0},
    // -xc               exclude COM ports (implicitly include LPT ports)
    { L"xc", OPT_FLAG_EXCLUDE_COM, OPT_FLAG_EXCLUDE_LPT },
    // -xl               exclude LPT/PRN ports (implicitly include COM ports)
    { L"xl", OPT_FLAG_EXCLUDE_LPT, OPT_FLAG_EXCLUDE_COM },
    // end of option list marker
    { NULL }
};

struct bus_match_info {
    const wchar_t*  buslabel;
    unsigned        flagnoparams;
    unsigned        minparams;
    unsigned        maxparams;
    unsigned        flagNparams[4];
    unsigned long   maxvalues[4];
    // probably a fn pointer here to customise handling for different buses
};

struct bus_match_info bus_list[] = {
    { L"USB", OPT_FLAG_USBMATCH_ANY, 1, 2, 
        { OPT_FLAG_USBMATCH_VID, OPT_FLAG_USBMATCH_PIDVID }, { 0xFFFF, 0xFFFF } },
    // end of list marker
    { NULL }
};

// get id arguments for vid or vid:pid
Bool getidarg(const wchar_t* arg, int params, unsigned long* pvid, unsigned long* ppid)
{
    const wchar_t* format = (params == 1) ? L"%4x" : L"%4x:%4x";

#if defined(OPTIONS_DEBUG)
    wprintf(L"swscanf returns %i\n", swscanf(argv[1], format, pvid, ppid));
#endif

    if (params != swscanf(arg, format, pvid, ppid)) {
        return False;
    }
    
#if defined(OPTIONS_DEBUG)
    wprintf( (params == 1) ? L"vid = 0x%04x\n" : L"vid = 0x%04x, pid = 0x%04x\n", *pvid, *ppid);
#endif

    // in range?
    if ( (*pvid > 0xFFFF) || (*ppid > 0xFFFF) ) {
        return False;
    }

    return True;
}


// returns number of arguments used
int matchoption(PortList* list, int argc, wchar_t** argv)
{
    unsigned long vid = 0, pid = 0; // temps for parsing vid & pid
    unsigned idx;
    const wchar_t* arg = *argv;

#if defined(OPTIONS_DEBUG)
    wprintf(L"arg = \"%s\", argc = %u\n", arg, argc); // debug aid
#endif

    if ( (*arg != L'-') && (*arg != L'/') ) {
        return 0;
    }

    arg++;

    // simple option flags
    for (idx = 0; opt_list[idx].opt_text != NULL; idx++) {
        if (!wcsicmp(arg, opt_list[idx].opt_text)) {
            list->opt_flags |= opt_list[idx].set_flags;
            list->opt_flags &= ~opt_list[idx].clear_flags;
            return 1;
        }   
    }

    // bus match option flags
    for (idx = 0; bus_list[idx].buslabel != NULL; idx++) {
        size_t strlen = wcslen(bus_list[idx].buslabel);

        if (!wcsnicmp(arg, bus_list[idx].buslabel, strlen)) {
            arg += strlen;
            if (*arg == L'\0') {
                list->opt_flags |= bus_list[idx].flagnoparams;
                return 1;
            } else if (*arg != L'=') {
                return 0;
            } else {
                // TODO make this more generic for other bus types & use bus_match_info
                arg += 1;
                if (wcschr(arg, L':')) {
                    // -usb=<vid>:<pid>  specify hex Vendor & Product IDs to match
                    if (!getidarg(arg, 2, &vid, &pid)) {
                        return 0;
                    }

                    // extend list memory
                    listadd(&(list->pidvidlist), (vid << 16) | pid);
                    list->opt_flags |= OPT_FLAG_USBMATCH_PIDVID;
                    return 1;
                } else {
                    // -usb=<vid>        specify a hex Vendor ID to match
                    if (!getidarg(arg, 1, &vid, &pid)) {
                        return 0;
                    }

                    // extend list memory
                    listadd(&(list->vidlist), vid);
                    list->opt_flags |= OPT_FLAG_USBMATCH_VID;
                    return 1;
                }
            }
        }   
    }


#if defined(SUPPORT_DEPRECATED_VID_OR_PID_OPTIONS)
    // options with parameters
    if (argc >= 2) {
        if (!wcsicmp(arg, L"pid")) {
            // -pid <vid>:<pid>  specify hex Vendor & Product IDs to match
            if (!getidarg(argv[1], 2, &vid, &pid)) {
                return 0;
            }

            // extend list memory
            listadd(&(list->pidvidlist), (vid << 16) | pid);
            list->opt_flags |= OPT_FLAG_USBMATCH_PIDVID;
            return 2;
        } else if (!wcsicmp(arg, L"vid")) {
            // -vid <vid>        specify a hex Vendor ID to match
            if (!getidarg(argv[1], 1, &vid, &pid)) {
                return 0;
            }

            // extend list memory
            listadd(&(list->vidlist), vid);
            list->opt_flags |= OPT_FLAG_USBMATCH_VID;
            return 2;
        }
    }
#endif

    // not recognised
    return 0;
}

Bool checkoptions(PortList* list, int argc, wchar_t** argv)
{
    while (argc) {
        int args_used = matchoption(list, argc, argv);

        if (args_used == 0) {
            // unrecognised / badly formatted option
            return False;
        }

        argc -= args_used;
        argv += args_used;
    } // while

    return True;
}


// boolean whether device matches pid, vid or vid:pid list entry
Bool findinlist(struct u32_list list, unsigned value)
{
    unsigned count = list.count;

    if (count > 0) {
        unsigned i;
        for (i = 0; i < count; i++) {
            if (list.ulist[i] == value) {
                return True; // match
            }
        }
    }
    return False;
}


Bool checkpidandvidlists(PortList* list, PortInfo* pInfo)
{
    Bool success = False;

    if (pInfo->isUSB) {
        // VID & PID seem valid enough to proceed with USB Id matching
        if (list->opt_flags & OPT_FLAG_USBMATCH_ANY) {
            success = True;
        }

        if (!success && list->opt_flags & OPT_FLAG_USBMATCH_VID) {
            success = findinlist(list->vidlist, pInfo->vendorId);
        }

        if (!success && (list->opt_flags & OPT_FLAG_USBMATCH_PIDVID)) {
            success = findinlist(list->pidvidlist, (pInfo->vendorId << 16) | pInfo->productId);
        }
    }
            
    return success;
}


void trygetdevice_regdword(HKEY devkey, wchar_t* keyname, DWORD* result, unsigned int* flags, unsigned int attribflag)
{
    if (devkey) {
        // Read in the name of the port
        DWORD sizeIn = sizeof(DWORD);
        DWORD sizeOut = sizeIn;
        DWORD type = 0;

        if (((RegQueryValueEx(devkey, keyname, NULL, &type, (LPBYTE)result, &sizeOut)) == ERROR_SUCCESS)
                    && (sizeOut == sizeIn) && (type == REG_DWORD)) {
            // record our success
            *flags |= attribflag;
        }
    }
}


wchar_t* getportname(HKEY devkey)
{
    //Read in the name of the port
    DWORD sizeIn = 0;
    DWORD sizeOut = sizeIn;
    DWORD type = 0;
    wchar_t* portname = NULL;

    // Win2k should return ERROR_MORE_DATA if buffer is too small, Win7 happily suceeds so need to check size too
    if (((RegQueryValueEx(devkey, L"PortName", NULL, &type, NULL, &sizeOut)) == ERROR_MORE_DATA)
                || (sizeOut > sizeIn)) {
        // check type
        if (REG_SZ != type) {
            errorprintf(L"expected PortName to be of type REG_SZ not %#X)", type);
        } else {
            // use sizeof(wchar_t) to workaround W2K returning size in characters instead of bytes
            wchar_t* buffer = calloc(sizeOut, sizeof(wchar_t));
            sizeIn = sizeOut;

            if (buffer && ((RegQueryValueEx(devkey, L"PortName", NULL, &type, (LPBYTE)buffer, &sizeOut)) == ERROR_SUCCESS)
                    && (sizeOut == sizeIn)) {
                // copy to new buffer that doesn't waste bytes on W2k workaround
                portname = wcs_dupsubstr(buffer, sizeIn);
            }

            free(buffer);
        }
    }

    return portname;
}


void getverboseportinfo(HKEY devkey, PortInfo* pInfo)
{
    // verbose details for legacy ports
    trygetdevice_regdword(devkey, L"PortAddress", &pInfo->portaddress, &pInfo->retrieved, RETRIEVED_PORTADDRESS);
    trygetdevice_regdword(devkey, L"Interrupt", &pInfo->interrupt, &pInfo->retrieved, RETRIEVED_INTERRUPT);

    // verbose details for multi-port devices
    trygetdevice_regdword(devkey, L"PortIndex", &pInfo->portindex, &pInfo->retrieved, RETRIEVED_PORTINDEX);
    trygetdevice_regdword(devkey, L"Indexed", &pInfo->indexed, &pInfo->retrieved, RETRIEVED_INDEXED);
}


PortInfo* getdevicesetupinfo(HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, unsigned opt_flags)
{
    PortInfo* pInfo = NULL;
    
    // get the registry key with the device's port settings
    HKEY devkey = SetupDiOpenDevRegKey(hDevInfo, pDeviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_QUERY_VALUE);

    if (devkey) {
        pInfo = (PortInfo*) calloc(1, sizeof(PortInfo));

        if (pInfo) {
            pInfo->portname = getportname(devkey);

            if (pInfo->portname) {
                if (opt_flags & OPT_FLAG_VERBOSE) {
                    getverboseportinfo(devkey, pInfo);
                }
            } else {
                // failed to get fullname
                free(pInfo);
                pInfo = NULL;
            }
        }

        // tidy up
        RegCloseKey(devkey);
    }

    return pInfo;
}


wchar_t* portstringproperty(HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, DWORD devprop)
{
    DWORD    type;
    DWORD    lastError;
    DWORD    buffersize = 0;
    wchar_t* strproperty = NULL;

    // first call gets property info, such as size & type
    SetupDiGetDeviceRegistryProperty(
        hDevInfo,
        pDeviceInfoData,
        devprop,
        &type,
        NULL,
        0,
        &buffersize);

    lastError = GetLastError();
    // continue if property currently defined for this port
    if (ERROR_INSUFFICIENT_BUFFER == lastError &&  ((REG_SZ == type) || (REG_MULTI_SZ == type))) {
        // sizeof(wchar_t) works around W2k MBCS bug per KB 888609. 
        wchar_t* buffer = calloc(buffersize, sizeof(wchar_t) );

        if (buffer && SetupDiGetDeviceRegistryProperty(
                hDevInfo,
                pDeviceInfoData,
                devprop,
                &type,
                (PBYTE)buffer,
                buffersize,
                NULL)) {

            // copy to new buffer that doesn't waste bytes on W2k workaround
            strproperty = wcs_dupsubstr(buffer, buffersize);
        }

        free(buffer);
    }

    if (!strproperty) {
        if ((ERROR_INVALID_DATA != lastError) && (ERROR_NO_SUCH_DEVINST != lastError)) {
            errorprintf(L"could not get property %#X - %#X)", devprop, lastError);
        }

        // check type
        if ((ERROR_SUCCESS == lastError) && (REG_SZ != type) && (REG_MULTI_SZ != type)) {
            errorprintf(L"expected property %#X to be of type REG_SZ or REG_MULTI_SZ not %#X)", devprop, type);
        }
    }

    return strproperty;
}


// duplicate a substring, max of length wide chars
// protects when copying registry strings that are not NIL terminated
wchar_t* wcs_dupsubstr(const wchar_t* string, size_t length)
{
    size_t alloclen; // = (string[length] == L'\0') ? length : (length + 1);
    wchar_t* buff = NULL;
    
    if (string != NULL) {
        for (alloclen = 0; (alloclen < length) && (string[alloclen] != L'\0'); alloclen++)
            ;

        if (alloclen > 0) {
            buff = calloc(alloclen + 1, sizeof(wchar_t));

            if (buff) {
                wcsncpy(buff, string, alloclen);
            }
        }
    }
    return buff;
}


// combine wcsstr() [= wcswcs()] and wcstoul()
Bool wcs_seekul(wchar_t** pString, const wchar_t* SubStr, unsigned long* pOutValue, int Radix)
{
    wchar_t* match = *pString;
    
    if (match && SubStr) {
        // find substring
        match = wcsstr(match, SubStr);
    }

    if (match) {
        // read following characters to get value
        match += wcslen(SubStr);
        *pOutValue = wcstoul(match, pString, Radix);
        if (match != *pString) {
            return True;
        }
    }
    return False;
}


/* Device Properties that we can pick from
 *  SPDRP_DEVICEDESC                  DeviceDesc (R/W)
 *  SPDRP_HARDWAREID                  HardwareID (R/W)
 *  SPDRP_COMPATIBLEIDS               CompatibleIDs (R/W)
 *  SPDRP_SERVICE                     Service (R/W)
 *  SPDRP_CLASS                       Class (R--tied to ClassGUID)
 *  SPDRP_CLASSGUID                   ClassGUID (R/W)
 *  SPDRP_DRIVER                      Driver (R/W)
 *  SPDRP_CONFIGFLAGS                 ConfigFlags (R/W)
 *  SPDRP_MFG                         Mfg (R/W)
 *  SPDRP_FRIENDLYNAME                FriendlyName (R/W)
 *  SPDRP_LOCATION_INFORMATION        LocationInformation (R/W)
 *  SPDRP_PHYSICAL_DEVICE_OBJECT_NAME PhysicalDeviceObjectName (R)
 *  SPDRP_CAPABILITIES                Capabilities (R)
 *  SPDRP_UI_NUMBER                   UiNumber (R)
 *  SPDRP_UPPERFILTERS                UpperFilters (R/W)
 *  SPDRP_LOWERFILTERS                LowerFilters (R/W)
 *  SPDRP_BUSTYPEGUID                 BusTypeGUID (R)
 *  SPDRP_LEGACYBUSTYPE               LegacyBusType (R)
 *  SPDRP_BUSNUMBER                   BusNumber (R)
 *  SPDRP_ENUMERATOR_NAME             Enumerator Name (R)
 *  SPDRP_SECURITY                    Security (R/W, binary form)
 *  SPDRP_SECURITY_SDS                Security (W, SDS form)
 *  SPDRP_DEVTYPE                     Device Type (R/W)
 *  SPDRP_EXCLUSIVE                   Device is exclusive-access (R/W)
 *  SPDRP_CHARACTERISTICS             Device Characteristics (R/W)
 *  SPDRP_ADDRESS                     Device Address (R)
 *  SPDRP_UI_NUMBER_DESC_FORMAT       UiNumberDescFormat (R/W)
 *  SPDRP_DEVICE_POWER_DATA           Device Power Data (R)
 *  SPDRP_REMOVAL_POLICY              Removal Policy (R)
 *  SPDRP_REMOVAL_POLICY_HW_DEFAULT   Hardware Removal Policy (R)
 *  SPDRP_REMOVAL_POLICY_OVERRIDE     Removal Policy Override (RW)
 *  SPDRP_INSTALL_STATE               Device Install State (R)
 *  SPDRP_LOCATION_PATHS              Device Location Paths (R)
 *  SPDRP_BASE_CONTAINERID            Base ContainerID (R)
 *  SPDRP_MAXIMUM_PROPERTY            Upper bound on ordinals
 */
Bool getportdetails(PortList* list, HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, PortInfo* pInfo)
{
    // get base information
    pInfo->friendlyname = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_FRIENDLYNAME);


#if defined(_DEBUG) && defined(DEBUG_DEV_PROPERTIES)
    //////////////////////////////////////////////////////////////////////////////
    // test code for peeking at device attribute 

    {
        wchar_t* tempstr = NULL;

        tempstr = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_ADDRESS);
        if (tempstr) {
            wprintf(L"SPDRP_ADDRESS = %s\n", tempstr);
            free(tempstr);
            tempstr = NULL;
        }

        tempstr = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_CHARACTERISTICS);
        if (tempstr) {
            wprintf(L"SPDRP_CHARACTERISTICS = %s\n", tempstr);
            free(tempstr);
            tempstr = NULL;
        }

        tempstr = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_UI_NUMBER);
        if (tempstr) {
            wprintf(L"SPDRP_UI_NUMBER = %s\n", tempstr);
            free(tempstr);
            tempstr = NULL;
        }

        tempstr = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_DRIVER);
        if (tempstr) {
            wprintf(L"SPDRP_DRIVER = %s\n", tempstr);
            free(tempstr);
            tempstr = NULL;
        }
    }
#endif

    //////////////////////////////////////////////////////////////////////////////////
    


    if (list->opt_flags & (OPT_FLAG_USBMATCH_SPECIFIED | OPT_FLAG_LONGFORM)) {
        // get Bus type, VID, PID & Revision
        pInfo->hardwareid = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_HARDWAREID);
        if (pInfo->hardwareid) {
            wchar_t* str = pInfo->hardwareid;
            size_t bus_len = wcsspn(str, L"ABCDEFGHIJKLMNOPQRSTUVWXYZ");

            // extract busname
            if (bus_len > 0) {
                pInfo->busname = wcs_dupsubstr(str, bus_len);
                str += bus_len;
            }

            if (wcs_seekul(&str, L"\\VID_", &(pInfo->vendorId), 16)) {
                pInfo->retrieved |= RETRIEVED_VENDOR_ID;

                if (wcs_seekul(&str, L"PID_", &(pInfo->productId), 16)) {
                    pInfo->retrieved |= RETRIEVED_PRODUCT_ID;
                    if ( (pInfo->vendorId < 0x10000) && (pInfo->productId < 0x10000) ) {
                        pInfo->isUSB = True;
                    }

                    if (wcs_seekul(&str, L"REV_", &(pInfo->revision), 16)) {
                        pInfo->retrieved |= RETRIEVED_REV;
                    }
                }
            }
        }
    }

    // Vendor / Manufacturer name
    if (list->opt_flags & OPT_FLAG_LONGFORM) {
        pInfo->product = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_DEVICEDESC);
        pInfo->vendor = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_MFG);

        // interesting values for verbose mode
        if (list->opt_flags & OPT_FLAG_VERBOSE) {
            pInfo->devclass = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_CLASS);
            pInfo->location = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_LOCATION_INFORMATION);

            /* TODO find & extract serial number
            if ( (*end == L'\\') && (list->opt_flags & OPT_FLAG_VERBOSE) ) {
                pInfo->serialnumber = wcsdup(end + 1);
            }	*/
        }
    }

    // for All or Verbose modes need the PhysDevObj, if set the device is available
    if (list->opt_flags & (OPT_FLAG_ALL | OPT_FLAG_VERBOSE)) {
        pInfo->physdevobj = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_PHYSICAL_DEVICE_OBJECT_NAME);
        if (pInfo->physdevobj) {
            pInfo->isAvailable = True;
        }
    }

    return True;
}


// compare portnames, eg COM5 and COM10 to create a stable sorted order
int portcmp(PortInfo* p1, PortInfo* p2)
{
    int res;
    if (p1->prefixlen != p2->prefixlen) {
        res = wcscmp(p1->portname, p2->portname);
    } else {
        res = wcsncmp(p1->portname, p2->portname, p1->prefixlen);
        if (res == 0)
            res = p1->portnumber - p2->portnumber;
    }

    return res;
}


Bool getdeviceinfo(PortList* list, HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData)
{
    Bool success = False;
    PortInfo* pInfo = getdevicesetupinfo(hDevInfo, pDeviceInfoData, list->opt_flags);

    if (pInfo) {
        // extract prefix and port number for port name sorting
        pInfo->prefixlen = wcscspn(pInfo->portname, L"0123456789");
        if (pInfo->prefixlen != wcslen(pInfo->portname)) {
            wchar_t* end;

            pInfo->portnumber = wcstoul(pInfo->portname + pInfo->prefixlen, &end, 10);
        }

        if ((list->opt_flags & (OPT_FLAG_EXCLUDE_COM | OPT_FLAG_EXCLUDE_LPT)) && (3 == pInfo->prefixlen)) {
            // use port name to distinguish COM & LPT ports
            Bool is_com_port = (0 == wcscmp(pInfo->portname, L"AUX")) ||
                (pInfo->portnumber && (0 == wcsncmp(pInfo->portname, L"COM", 3)));

            if (list->opt_flags & OPT_FLAG_EXCLUDE_COM) {
                // exclude AUX & COM ports
                if (!is_com_port) {
                    success = getportdetails(list, hDevInfo, pDeviceInfoData, pInfo);
                }
            } else { // OPT_FLAG_EXCLUDE_LPT - only AUX & COM ports
                if (is_com_port) {
                    success = getportdetails(list, hDevInfo, pDeviceInfoData, pInfo);
                }
            }
        } else {
            success = getportdetails(list, hDevInfo, pDeviceInfoData, pInfo);
        }

        if (success && (list->opt_flags & OPT_FLAG_EXCLUDE_AVAILABLE) && pInfo->isAvailable) {
            // device is available, but we've been requested to exclude available this time!
            success = False;
        }

        if (success && (list->opt_flags & OPT_FLAG_USBMATCH_SPECIFIED)) {
            success = checkpidandvidlists(list, pInfo);
        }

        if (success) {
            // sort by name
            
            if ((list->ports == NULL) || (portcmp(pInfo, list->ports) < 0)) {
                // place at front of linked list
                pInfo->next = list->ports;
                list->ports = pInfo;
            } else {
                PortInfo* prev = list->ports;
                PortInfo* next = prev->next;
                while (next && (portcmp(pInfo, next) > 0)) {
                    prev = next;
                    next = next->next;
                }
                pInfo->next = next;
                prev->next = pInfo;
            }
        } else {
            // cleanup allocated strings & memory
            free(pInfo->friendlyname);
            free(pInfo->busname);
            free(pInfo->product);
            free(pInfo->vendor);
            free(pInfo->portname);
            free(pInfo->hardwareid);
            free(pInfo->location);
            free(pInfo->physdevobj);
            free(pInfo->devclass);
#if defined(SUPPORT_SERIAL_NUMBER_REPORTING)
            free(pInfo->serialnumber);
#endif

            free(pInfo);
        }
    }

    return success;
}


unsigned listdevices(PortList* list, HDEVINFO hDevInfo)
{
    DWORD lastError = 0;
    SP_DEVINFO_DATA DeviceInfoData;
    DWORD dev;
    unsigned portcount = 0;

    ZeroMemory(&DeviceInfoData, sizeof(SP_DEVINFO_DATA));
    DeviceInfoData.cbSize = sizeof(SP_DEVINFO_DATA);

    for (dev = 0; SetupDiEnumDeviceInfo(hDevInfo, dev, &DeviceInfoData); dev++)
    {
        if (getdeviceinfo(list, hDevInfo, &DeviceInfoData)) {
            portcount ++;
        }
    }

    lastError = GetLastError();
    if (NO_ERROR !=lastError && ERROR_NO_MORE_ITEMS != lastError) {
        // TODO is there anything else we can do here for error handling?
        errorprintf(L"unexpected error fetching Device Info - 0x%X", lastError);
    }

    return portcount;
}


unsigned listclass(PortList* list, const GUID *guid)
{
    /*
       device interface GUIDs of interest to us include:
       GUID_DEVINTERFACE_COMPORT - internal serial port
       GUID_DEVINTERFACE_PARALLEL - Parallel port
    */ 
    HDEVINFO hDevInfo = INVALID_HANDLE_VALUE;
    unsigned count = 0;
    DWORD    devflags = (list->opt_flags & OPT_FLAG_ALL) ? 0 : DIGCF_PRESENT;

    /* Create a HDEVINFO with devices matching GUID & user choise of -a or -p
     * MSDN example code I've seen for this API includes DIGCF_DEVICEINTERFACE,
     * but for me this stops any COM ports from being found.
     */
    hDevInfo = SetupDiGetClassDevs(guid, 0, 0, devflags);

    if (hDevInfo == INVALID_HANDLE_VALUE)
    {
        // unrecoverable error
        wchar_t guid_string[40];
        swprintf(guid_string, sizeof(guid_string) / sizeof(wchar_t),
            L"{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}",
            guid->Data1, guid->Data2, guid->Data3,
            guid->Data4[0], guid->Data4[1], guid->Data4[2],
            guid->Data4[3], guid->Data4[4], guid->Data4[5],
            guid->Data4[6], guid->Data4[7]);

        errorprintf(L"error calling SetupDiGetClassDevs with %s - 0x%X", guid_string, GetLastError());
        exit(-1);
    } else {
        // Enumerate through all devices in Set
        count = listdevices(list, hDevInfo);
        
        //  Cleanup
        SetupDiDestroyDeviceInfoList(hDevInfo);
    }

    return count;
}


void listports(PortList* list)
{
    /* device setup GUIDs to look for are:
       GUID_DEVCLASS_PORTS single COM / LPT ports
       GUID_DEVCLASS_MODEM modem ports are not included in GUID_DEVCLASS_PORTS
       GUID_DEVCLASS_MULTIPORTSERIAL multiple COM ports on single (PCI) card
    */

    PortInfo*       p;
    // get info about ports
    unsigned count = listclass(list, &GUID_DEVCLASS_PORTS);

    // add modems & multiport serial ports, unless COM ports are excluded
    if ((list->opt_flags & OPT_FLAG_EXCLUDE_COM) == 0) {
        count += listclass(list, &GUID_DEVCLASS_MODEM);
        count += listclass(list, &GUID_DEVCLASS_MULTIPORTSERIAL);
    }

    // print details of all the (matching) ports we found
    p = list->ports;       // linked list of port info

    if (list->opt_flags & OPT_FLAG_LONGFORM) {
        wprintf(L"Port   %sBus    VID  PID  Rev  Friendly name\n",
            list->opt_flags & (OPT_FLAG_ALL | OPT_FLAG_VERBOSE) ? L"A " : L"");

        for (; p; p = p->next) {
            wprintf(L"%-6s ", p->portname);

            // device availability only for Verbose or All listings
            if (list->opt_flags & (OPT_FLAG_ALL | OPT_FLAG_VERBOSE) ) {
                wprintf(p->isAvailable ? L"A " : L". ");
            }

            if (p->isUSB) {
                const wchar_t* fmt_4hex = L"%04X ";
                const wchar_t* spaces5  = L"     ";

                wprintf(L"%-6s ", p->busname ? p->busname : L" ");

                // USB Vendor Id & Product Id only retrieved together
                wprintf(fmt_4hex, p->vendorId);
                wprintf(fmt_4hex, p->productId);
                wprintf(p->retrieved & RETRIEVED_REV ? fmt_4hex : spaces5, p->revision);
            } else {
                // if no VID we have more space for long busnames
                wprintf(L"%-21s ", p->busname ? p->busname : L" ");
            }

            if (p->friendlyname) {
                wprintf(L"%s\n", p->friendlyname);
            } else {
                wprintf(L"\n");
            }

            // extra info for verbose mode
            if (list->opt_flags & OPT_FLAG_VERBOSE) {
                wchar_t* indent = L"         ";

                if (p->vendor) {
                    wprintf(L"%sVendor: %s\n", indent, p->vendor);
                }
                if (p->product) {
                    wprintf(L"%sProduct: %s\n", indent, p->product);
                }
                if (p->devclass) {
                    wprintf(L"%sDevice Class: %s\n", indent, p->devclass);
                }
                if (p->hardwareid) {
                    wprintf(L"%sHardware Id: %s\n", indent, p->hardwareid);
                }
#if defined(SUPPORT_SERIAL_NUMBER_REPORTING)
                if (p->serialnumber) {
                    wprintf(L"%sSerial number: %s\n", indent, p->serialnumber);
                }
#endif
                if (p->physdevobj) {
                    wprintf(L"%sPhysical Device Object: %s\n", indent, p->physdevobj);
                }
                if (p->location) {
                    wprintf(L"%sLocation Info: %s\n", indent, p->location);
                }

                // ISA legacy hardware port
                if ((p->retrieved & (RETRIEVED_PORTADDRESS | RETRIEVED_INTERRUPT)) == (RETRIEVED_PORTADDRESS | RETRIEVED_INTERRUPT)) {
                    wprintf(L"%sLegacy port -- address %04lX, interrupt %lu\n", indent, p->portaddress, p->interrupt);
                }

                // multiport device
                if ((p->retrieved & (RETRIEVED_PORTINDEX | RETRIEVED_INDEXED)) == (RETRIEVED_PORTINDEX | RETRIEVED_INDEXED)) {
                    wprintf(L"%sMulti-port device -- port ", indent);
                        
                    wprintf(p->indexed ? L"index %lu\n" : L"%bitmap 0x%04lX\n", p->portindex);
                }

                // if there is another port to print add a spacing line
                if (p->next) {
                    wprintf(L"\n");
                }
            }
        }
    } else {
        wprintf(L"Port   %sFriendly name\n", list->opt_flags & OPT_FLAG_ALL ? L"A " : L"");

        for (; p; p = p->next) {
            wprintf(L"%-6s ", p->portname);

            if (list->opt_flags & OPT_FLAG_ALL) {
                wprintf(p->isAvailable ? L"A " : L". ");
            }

            if (p->friendlyname) {
                wprintf(L"%s\n", p->friendlyname);
            } else {
                wprintf(L"\n");
            }
        }
    }

    wprintf(L"\n%u %sport%s found.\n", count, 
        (list->opt_flags & OPT_FLAG_USBMATCH_SPECIFIED) ? L"matching " : L"",
        (count != 1) ? L"s" : L"");
}


// Unicode argv[] version of main()
int wmain(int argc, wchar_t* argv[])
{
    PortList list;

    memset(&list, 0, sizeof(PortList));

#ifdef _DEBUG
    // debug build: default to test maximum number of details
    list.opt_flags = OPT_FLAG_ALL | OPT_FLAG_LONGFORM | OPT_FLAG_VERBOSE;
    // list.opt_flags = OPT_FLAG_HELP;
    // list.opt_flags = OPT_FLAG_HELP_COPYRIGHT
    
    {
        wchar_t* testargs[] = { L"/usb=04d8:000A", L"-usb=421" };
        checkoptions(&list, 2, testargs);
    }
    
#endif

    if (argc > 1) {
        // parse command line
        if (!checkoptions(&list, argc - 1, argv + 1)) {
            errorprint(L"Bad parameter");

            // help text & exit
            usage(False, False);
            return -1;
        }
    }

    if (list.opt_flags & (OPT_FLAG_HELP | OPT_FLAG_HELP_COPYRIGHT)) {
        // verbose help and or copyright text
        usage(list.opt_flags & OPT_FLAG_HELP, list.opt_flags & OPT_FLAG_HELP_COPYRIGHT); 
    } else {
        // make & print port list
        listports(&list);
    }

    return 0;
}

