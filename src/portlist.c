/*
    portlist.c - list COM (serial) and/or LPT (parallel) ports on MS Windows 

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
//#define OPTIONS_DEBUG 
#endif


typedef unsigned Bool;
const unsigned False = 0;
const unsigned True = 1;

// English messages for help & bad options
#define program_name L"portlist"
const wchar_t* name_msg = program_name;
const wchar_t* version_msg = L"0.9";
const wchar_t* copyright_msg = L"Copyright (c) 2013, 2014 Anthony Naggs";

const wchar_t* free_sw_msg =
    program_name L" comes with ABSOLUTELY NO WARRANTY.\n"
    L"This software is free, you are welcome to redistribute it under certain conditions.\n"
    L"Type `" program_name L" -c' for Copyright & Warranty details.\n"
    program_name L"source and binary files are available from https://github.com/tonynaggs/portlist\n\n";

const wchar_t* long_copyright_msg =
    L"Limited assignment of rights through the GNU General Public License version 2\n"
    L"is described below:\n"
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
    
const wchar_t* usage_msg = 
    L"Usage: " program_name L" [-a] [-c] [-h|-?] [-l] [-pid <vid>:<pid>] [-vid <vid>] [-v] [-x] [-xc] [-xl]\n"
    L"\t-a                list available (default) and also remembered ports\n"
    L"\t-c                show Copyright and Warranty details (GNU General Public License v2)\n"
    L"\t-h or -?          show this help text plus examples\n"
    L"\t-l                long including Bus type, Vendor & Product IDs\n"
    L"\t-vid <vid>        specify a hex Vendor ID to match\n"
    L"\t-pid <vid>:<pid>  specify pair of hex Vendor & Product IDs to match\n"
    L"\t-v                verbose multi-line per port list (implies -l)\n"
    L"\t-x                exclude available ports => list only remembered ports\n"
    L"\t-xc               exclude COM ports\n"
    L"\t-xl               exclude LPT/PRN ports\n"
    L"Note that multiple '-vid' and/or '-pid' parameters can be specified.\n"
    L"Options can start with / or - and be upper or lowercase.\n\n";

const wchar_t* examples_msg = 
    L"Examples:\n"
    L"\t" program_name L"                    : list available ports and description\n"
    L"\t" program_name L" -l                 : longer, detailed list of available ports\n"
    L"\t" program_name L" -a                 : all available & remembered ports\n"
    L"\t" program_name L" /XL                : exclude printer ports => COM ports only\n"
    L"\t" program_name L" -pid 2341:0001     : match Arduino Uno VID/PID\n"
    L"\t" program_name L" /pid 04d8:000A     : match Microchip USB serial port ref\n"
    L"\t" program_name L" -pid 1d50:6098     : match Aperture Labs' RFIDler\n"
    L"\t" program_name L" /vid 0403          : match FTDI's Vendor ID (to find USB serial bridge chips)\n"
    L"\t" program_name L" -vid 4e8 -vid 421  : match either Samsung or Nokia VIDs\n\n";

// bit flags for option switches
#define OPT_FLAG_ALL                0x00000001
#define OPT_FLAG_LONGFORM           0x00000002
#define OPT_FLAG_VERBOSE            0x00000004
#define OPT_FLAG_MATCH_VID          0x00000010
#define OPT_FLAG_MATCH_PIDVID       0x00000020
#define OPT_FLAG_EXCLUDE_COM        0x00000100
#define OPT_FLAG_EXCLUDE_LPT        0x00000200
#define OPT_FLAG_EXCLUDE_AVAILABLE  0x00000400

#define OPT_FLAG_HELP_COPYRIGHT     0x40000000
#define OPT_FLAG_HELP               0x80000000

// if any of the PID or VID matching options are specified
#define OPT_FLAG_MATCH_SPECIFIED    (OPT_FLAG_MATCH_PIDVID | OPT_FLAG_MATCH_VID)


// bit flags for retrieved data
#define RETRIEVED_VID               0x00000001
#define RETRIEVED_PID               0x00000002
#define RETRIEVED_REV               0x00000004
#define RETRIEVED_PORTADDRESS       0x00000008
#define RETRIEVED_INTERRUPT         0x00000010
#define RETRIEVED_PORTINDEX         0x00000020
#define RETRIEVED_INDEXED           0x00000040


#define BUSNAME_MAX                 9      /* long enough for "Bluetooth" */

////////////////////////////////////////////////
// struct definintions
////////////////////////////////////////////////

struct u32_list {
    unsigned*   ulist;
    unsigned    count;
    unsigned    max;
};


typedef struct portinfo {
    wchar_t*            name;           // COM1, PRN, ...
    wchar_t*            description;    // product description eg "USB Serial Port"

    // sorting key info
    size_t              prefixlen;      // length of "COM", "LPT" prefix, or strlen of name
    unsigned            portnumber;     // upto 3 digit number following COM or LPT

    // optional info for long listing
    wchar_t*            vendor;
    wchar_t             busname[BUSNAME_MAX + 1];     // 
    unsigned            vid;
    unsigned            pid;
    unsigned            revision;
    unsigned            retrieved;

    // optional info for verbose listing
    wchar_t*            hardwareid;
    wchar_t*            location;
    wchar_t*            physdevobj;
    wchar_t*            devclass;
    Bool                is_available;

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
Bool checkpidandvidlists(PortList* list, unsigned vid, unsigned pid);
void trygetdevice_regdword(HKEY devkey, wchar_t* keyname, DWORD* result, unsigned int* flags, unsigned int attribflag);
wchar_t* getportname(HKEY devkey);
void getverboseportinfo(HKEY devkey, PortInfo* pInfo);
PortInfo* getdevicesetupinfo(HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, unsigned opt_flags);
wchar_t* portstringproperty(HDEVINFO hDevInfo, SP_DEVINFO_DATA* pDeviceInfoData, DWORD devprop);
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
    return fwprintf(stderr, L"%s: %s\n", name_msg, message);
}


// print program name & formatted error message to stderr
int errorprintf(const wchar_t* format, ...)
{
    int res;
    va_list arglist;

    va_start(arglist, format);

    res = fwprintf(stderr, L"%s: ", name_msg);

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
        name_msg, version_msg, copyright_msg);

    if (help_copyright) {
        fputws(long_copyright_msg, stderr);
    } else {
        fputws(free_sw_msg, stderr);
    }

    if (help_examples || !help_copyright) {
        fputws(usage_msg, stderr);
    }

    if (help_examples) {
        fputws(examples_msg, stderr);
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


// get id arguments for vid or vid:pid
Bool getidarg(int argc, wchar_t** argv, int params, unsigned long* pvid, unsigned long* ppid)
{
    const wchar_t* format = (params == 1) ? L"%4x" : L"%4x:%4x";

    if (argc < 2) {
        return False;
    }

#if defined(OPTIONS_DEBUG)
    wprintf(L"swscanf returns %i\n", swscanf(argv[1], format, pvid, ppid));
#endif

    if (params != swscanf(argv[1], format, pvid, ppid)) {
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

    // options with parameters
    if (!wcsicmp(arg, L"pid")) {
        // -pid <vid>:<pid>  specify hex Vendor & Product IDs to match
        if (!getidarg(argc, argv, 2, &vid, &pid)) {
            return 0;
        }

        // extend list memory
        listadd(&(list->pidvidlist), (vid << 16) | pid);
        list->opt_flags |= OPT_FLAG_MATCH_PIDVID;
        return 2;
    } else if (!wcsicmp(arg, L"vid")) {
        // -vid <vid>        specify a hex Vendor ID to match
        if (!getidarg(argc, argv, 1, &vid, &pid)) {
            return 0;
        }

        // extend list memory
        listadd(&(list->vidlist), vid);
        list->opt_flags |= OPT_FLAG_MATCH_VID;
        return 2;
    }

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


Bool checkpidandvidlists(PortList* list, unsigned vid, unsigned pid)
{
    return findinlist(list->vidlist, vid) || findinlist(list->pidvidlist, (vid << 16) | pid);
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

    // Win2k should return ERROR_MORE_DATA if buffer is too small, Win7 happily suceeds so need to check size too
    if (((RegQueryValueEx(devkey, L"PortName", NULL, &type, NULL, &sizeOut)) == ERROR_MORE_DATA)
                || (sizeOut > sizeIn)) {
        // check type
        if (REG_SZ != type) {
            errorprintf(L"expected PortName to be of type REG_SZ not %#X)", type);
        } else {
            // use calloc & +1 to workaround issue that RegQueryValueEx does not ensure NIL terminator on string
            // use sizeof(wchar_t) to workaround W2K returning sie in characters instead of bytes
            wchar_t* buffer = calloc(sizeOut + 1, sizeof(wchar_t));
            sizeIn = sizeOut;

            if (buffer && ((RegQueryValueEx(devkey, L"PortName", NULL, &type, (LPBYTE)buffer, &sizeOut)) == ERROR_SUCCESS)
                    && (sizeOut == sizeIn)) {
                wchar_t* buff2;

                // copy to new buffer that doesn't waste bytes on W2k workaround
                buff2 = wcsdup(buffer);
                if (buff2) {
                    free(buffer);
                    return buff2;
                }
            }

            free(buffer);
        }
    }

    return NULL;
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
            pInfo->name = getportname(devkey);

            if (pInfo->name) {
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
        // Allocate a buffer, calloc & +1 ensure string is NIL terminated
        // sizeof(wchar_t) works around W2k MBCS bug per KB 888609. 
        wchar_t* buffer = calloc(buffersize + 1, sizeof(wchar_t) );

        if (buffer && SetupDiGetDeviceRegistryProperty(
                hDevInfo,
                pDeviceInfoData,
                devprop,
                &type,
                (PBYTE)buffer,
                buffersize,
                NULL)) {
            wchar_t* buff2;

            // copy to new buffer that doesn't waste bytes on W2k workaround
            buff2 = wcsdup(buffer);
            if (buff2) {
                free(buffer);
                return buff2;
            }
        }

        free(buffer);
    }

    if ((ERROR_INVALID_DATA != lastError) && (ERROR_NO_SUCH_DEVINST != lastError)) {
        errorprintf(L"could not get property %#X - %#X)", devprop, lastError);
    }

    // check type
    if ((ERROR_SUCCESS == lastError) && (REG_SZ != type) && (REG_MULTI_SZ != type)) {
        errorprintf(L"expected property %#X to be of type REG_SZ or REG_MULTI_SZ not %#X)", devprop, type);
    }

    return NULL;
}


/* Device Properties that we can pick from
 *  SPDRP_DEVICEDESC                  DeviceDesc (R/W
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
    pInfo->description = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_DEVICEDESC);
    if (NULL == pInfo->description) {
        return False;
    }

    if (list->opt_flags & (OPT_FLAG_MATCH_SPECIFIED | OPT_FLAG_LONGFORM)) {
        // get VID, PID & Revision
        pInfo->hardwareid = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_HARDWAREID);
        if (pInfo->hardwareid) {
            int i = 0;
            wchar_t* str = pInfo->hardwareid;
            wchar_t* end = str;

            while (isupper(str[i]) && (i < BUSNAME_MAX)) {
                pInfo->busname[i] = str[i];
                i++;
            }

            str = wcswcs(str + i, L"VID_");
            if (str) {
                pInfo->vid = wcstoul(str + 4, &end, 16);
                pInfo->retrieved |= RETRIEVED_VID;

                str = wcswcs(end, L"PID_");
                if (str) {
                    pInfo->pid = wcstoul(str + 4, &end, 16);
                    pInfo->retrieved |= RETRIEVED_PID;

                    str = wcswcs(end, L"REV_");
                    if (str) {
                        pInfo->revision = wcstoul(str + 4, &end, 16);
                        pInfo->retrieved |= RETRIEVED_REV;
                    }
                }
            }
        }
    }

    // Vendor / Manufacturer name
    if (list->opt_flags & OPT_FLAG_LONGFORM) {
        pInfo->vendor = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_MFG);

        // interesting values for verbose mode
        if (list->opt_flags & OPT_FLAG_VERBOSE) {
            pInfo->devclass = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_CLASS);
            pInfo->location = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_LOCATION_INFORMATION);
        }
    }

    // for All or Verbose modes need the PhysDevObj, if set the device is available
    if (list->opt_flags & (OPT_FLAG_ALL | OPT_FLAG_VERBOSE)) {
        pInfo->physdevobj = portstringproperty(hDevInfo, pDeviceInfoData, SPDRP_PHYSICAL_DEVICE_OBJECT_NAME);
        if (pInfo->physdevobj) {
            pInfo->is_available = True;
        }
    }

    return True;
}


// compare portnames, eg COM5 and COM10 to create a stable sorted order
int portcmp(PortInfo* p1, PortInfo* p2)
{
    int res;
    if (p1->prefixlen != p2->prefixlen) {
        res = wcscmp(p1->name, p2->name);
    } else {
        res = wcsncmp(p1->name, p2->name, p1->prefixlen);
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
        // extract info for name sorting
        pInfo->prefixlen = wcscspn(pInfo->name, L"0123456789");
        if (pInfo->prefixlen != wcslen(pInfo->name)) {
            wchar_t* end;

            pInfo->portnumber = wcstoul(pInfo->name + pInfo->prefixlen, &end, 10);
        }

        if ((list->opt_flags & (OPT_FLAG_EXCLUDE_COM | OPT_FLAG_EXCLUDE_LPT)) && (3 == pInfo->prefixlen)) {
            // use port name to distinguish COM & LPT ports
            Bool is_com_port = (0 == wcscmp(pInfo->name, L"AUX")) || (pInfo->portnumber && (0 == wcsncmp(pInfo->name, L"COM", 3)));

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

        if (success && (list->opt_flags & OPT_FLAG_EXCLUDE_AVAILABLE) && pInfo->is_available) {
            // device is available, but we've been requested to exclude available this time!
            success = False;
        }

        if (success && (list->opt_flags & OPT_FLAG_MATCH_SPECIFIED)) {
            if ((list->opt_flags & (OPT_FLAG_MATCH_VID | OPT_FLAG_MATCH_PIDVID)) && !(pInfo->retrieved & RETRIEVED_VID)) {
                // no VID retrieved, so cannot match
                success = False;
            } else if ((list->opt_flags & OPT_FLAG_MATCH_PIDVID) && !(pInfo->retrieved & RETRIEVED_PID)) {
                // no PID retrieved, so cannot match
                success = False;
            } else {
                success = checkpidandvidlists(list, pInfo->vid, pInfo->pid);
            }
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
            // cleanup any allocated strings
            free(pInfo->description);
            free(pInfo->vendor);
            free(pInfo->name);
            free(pInfo->hardwareid);
            free(pInfo->location);
            free(pInfo->physdevobj);
            free(pInfo->devclass);

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
       GUID_DEVCLASS_MULTIPORTSERIAL multiple COM ports
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
        wprintf(L"Port    A Bus    VID  PID  Rev  Product, Vendor\n");

        while (p) {
            wprintf(L"%-6s  ", p->name);

            // physdevobj string is set if device is available
            wprintf(L"%c ", p->is_available ? 'Y' : 'n');

            if (p->retrieved & (RETRIEVED_VID | RETRIEVED_PID | RETRIEVED_REV)) {
                const wchar_t* fmt_4hex = L"%04X ";
                const wchar_t* spaces5  = L"     ";

                wprintf(L"%-6s ", p->busname);
                wprintf(p->retrieved & RETRIEVED_VID ? fmt_4hex : spaces5, p->vid);
                wprintf(p->retrieved & RETRIEVED_PID ? fmt_4hex : spaces5, p->pid);
                wprintf(p->retrieved & RETRIEVED_REV ? fmt_4hex : spaces5, p->revision);
            } else {
                // output is prettier for long busnames (e.g. if "Bluetooth" is ever found)
                wprintf(L"%-21s ", p->busname);
            }
            wprintf(L"%.30s, %.20s\n", p->description, p->vendor);

            if (list->opt_flags & OPT_FLAG_VERBOSE) {
                wchar_t* indent = L"\t  ";

                if (p->devclass) {
                    wprintf(L"%sDevice Class: %s\n", indent, p->devclass);
                }
                if (p->hardwareid) {
                    wprintf(L"%sHardware Id: %s\n", indent, p->hardwareid);
                }
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
                    wprintf(p->indexed ? L"%s%s index %lu\n" : L"%s%s bitmap 0x%04lX\n", indent,
                        L"Multi-port device -- port ", p->portindex);
                }

                // if there is another port to print add a spacing line
                if (p->next) {
                    wprintf(L"\n");
                }
            }

            p = p->next;
        }
    } else {
        wprintf(L"Port   A Description\n");

        while (p) {
            wprintf(L"%-6s %c %.30s\n", p->name, p->is_available ? 'Y' : 'n', p->description);

            p = p->next;
        }
    }

    wprintf(L"\n%u %sport%s found.\n", count, 
        (list->opt_flags & OPT_FLAG_MATCH_SPECIFIED) ? L"matching " : L"",
        (count != 1) ? L"s" : L"");
}


// Unicode argv[] version of main()
int wmain(int argc, wchar_t* argv[])
{
    PortList list;

    memset(&list, 0, sizeof(PortList));

#ifdef _DEBUG
    // debug build: default to test maximum number of details
    list.opt_flags |= OPT_FLAG_ALL | OPT_FLAG_LONGFORM | OPT_FLAG_VERBOSE;
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

