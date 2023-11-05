//=====================================================================================================================================================================================================
//
// SpongySoft ShortCut
// Copyright (C) SpongySoft. All rights reserved.
//
//=====================================================================================================================================================================================================
#include "stdafx.h"
#include "shortcut.h"


//=====================================================================================================================================================================================================
// Define some version stuff, and other definitions and macros
//=====================================================================================================================================================================================================
#define THIS_VERSION "3.19.231015"
#define MAX_STRING 16384

#ifdef _DEBUG
	#define FLAVOR "(DEBUG) "
#else
    #define FLAVOR ""
#endif

// array size...
template<typename T, size_t size> size_t inline ARRAY_SIZE(const T(&)[size]) { return size; }


//=====================================================================================================================================================================================================
// namespaces in use
//=====================================================================================================================================================================================================
using namespace SpongySoft;



//=====================================================================================================================================================================================================
// Display the usage
//=====================================================================================================================================================================================================
void usage(int option=0)
{
	printf("Usage:\n\n");
	printf("   shortcut.exe [create | edit | query] shortcutfile [options]\n\n");
	printf("Options:\n");
	printf("   -t   target\n");
	printf("   -w   working folder\n");
	printf("   -i   iconfile iconindex\n");
	printf("   -d   description\n");
	printf("   -a   arguments (must be the last option)\n");
	printf("   -c   index color\n");
	printf("   -nl  no logo\n");
	printf("   -ws  window size (x y)\n");
	printf("   -bs  buffer size (x y)\n");
	printf("   -fa  fill attribute (fore back)\n");
	printf("   -pf  popup fill attribute (fore back)\n");
	printf("   -qe  num (quick edit, 0 for false, 1 for true)\n");
	printf("   -im  num (insert mode, 0 for false, 1 for true)\n");
	printf("   -ra  (0 or 1) run as administrator\n");
	printf("   -lw  (0 or 1) line wrapping (for v2 shortcuts)\n");
	printf("   -v2  (0 or 1) force v2 (for v2 shortcuts)\n");
	printf("   -fn  {name} font name\n");
	printf("   -ff  {num} font family\n");
	printf("   -fw  {num} font weight\n");
	printf("   -fs  {num} {num} font size\n");
	printf("   -min start minimized\n");
}

//	Props.wFillAttribute:          7
//	Props.wPopupFillAttribute:     245
//	Props.dwScreenBufferSize:      160,9999
//	Props.dwWindowSize:            160,75
//	Props.dwWindowOrigin:          0,0
//	Props.nFont:                   0
//	Props.nInputBufferSize:        0
//	Props.dwFontSize:              8,12
//	Props.uFontFamily:             48
//	Props.uFontWeight:             400
//	Props.FaceName:                Terminal
//	Props.uCursorSize:             25
//	Props.bFullScreen:             0
//	Props.bQuickEdit:              1
//	Props.bInsertMode:             1
//	Props.bAutoPosition:           1
//	Props.uHistoryBufferSize:      50
//	Props.uNumberOfHistoryBuffers: 4
//	Props.bHistoryNoDup:           0
//	Props.ColorTable[ 0]:          00000000
//	Props.ColorTable[ 1]:          00800000
//	Props.ColorTable[ 2]:          00008000
//	Props.ColorTable[ 3]:          00808000
//	Props.ColorTable[ 4]:          00000080
//	Props.ColorTable[ 5]:          00800080
//	Props.ColorTable[ 6]:          00008080
//	Props.ColorTable[ 7]:          00C0C0C0
//	Props.ColorTable[ 8]:          00808080
//	Props.ColorTable[ 9]:          00FF0000
//	Props.ColorTable[10]:          0000FF00
//	Props.ColorTable[11]:          00FFFF00
//	Props.ColorTable[12]:          000000FF
//	Props.ColorTable[13]:          00FF00FF
//	Props.ColorTable[14]:          0000FFFF
//	Props.ColorTable[15]:          00FFFFFF


//=====================================================================================================================================================================================================
// Set the default console props to use
//=====================================================================================================================================================================================================
static NT_CONSOLE_PROPS defaultProps =
{
	{
		sizeof(defaultProps),
		NT_CONSOLE_PROPS_SIG,
	},

	7,						// wFillAttribute
	245,					// wPopupFillAttribute
	160,9999,				// dwScreenBufferSize
	160,75,					// dwWindowSize
	0,0,					// dwWindowOrigin
	0,						// nFont
	0,						// nInputBufferSize
	8,12,					// dwFontSize
	48,						// uFontFamily
	400,					// uFontWeight
	{ _T("Terminal") },		// FaceName
	25,						// uCursorSize
	0,						// bFullScreen
	1,						// bQuickEdit
	1,						// bInsertMode
	1,						// bAutoPosition
	50,						// uHistoryBufferSize
	4,						// uNumberOfHistoryBuffers
	0,						// bHistoryNoDup
							//
	0x00000000,				// ColorTable[ 0]
	0x00800000,				// ColorTable[ 1]
	0x00008000,				// ColorTable[ 2]
	0x00808000,				// ColorTable[ 3]
	0x00000080,				// ColorTable[ 4]
	0x00800080,				// ColorTable[ 5]
	0x00008080,				// ColorTable[ 6]
	0x00C0C0C0,				// ColorTable[ 7]
	0x00808080,				// ColorTable[ 8]
	0x00FF0000,				// ColorTable[ 9]
	0x0000FF00,				// ColorTable[10]
	0x00FFFF00,				// ColorTable[11]
	0x000000FF,				// ColorTable[12]
	0x00FF00FF,				// ColorTable[13]
	0x0000FFFF,				// ColorTable[14]
	0x00FFFFFF,				// ColorTable[15]

};


//=====================================================================================================================================================================================================
// Hex conversion
//=====================================================================================================================================================================================================
#define HEXVAL(c) ((((c)>='0')&&((c)<='9')) ? (c)-'0' :((((c)>='A')&&((c)<='F')) ? (c)+10-'A' :((((c)>='a')&&((c)<='f')) ? (c)+10-'a' : 0)))

DWORD _htoi(const TCHAR *s)
{
	DWORD d=0;
	while (s && *s)
	{
		d = 16*d + HEXVAL(*s);
		s++;
	}

	return d;
}


//=====================================================================================================================================================================================================
// Convert a console option to a string
//=====================================================================================================================================================================================================
char *ConsoleOptionToString(ShortCut::v2ConsoleOption value)
{
	switch (value)
	{
		#define SWITCHCASE(o) case ShortCut::v2ConsoleOption::o: return #o;break;
		SWITCHCASE(forcev2);
		SWITCHCASE(wraptext);
		SWITCHCASE(filteronpaste);
		SWITCHCASE(ctrlkeysdisabled);
		SWITCHCASE(lineselection);
		SWITCHCASE(windowtransparency);
		SWITCHCASE(trimzeros);
		#undef SWITCHCASE
		default: return "unknown";
	}
}


//=====================================================================================================================================================================================================
// Convert bool to string
//=====================================================================================================================================================================================================
char *ConsoleBoolToString(ShortCut::v2ConsoleBool value)
{
	switch (value)
	{
		#define SWITCHCASE(o) case ShortCut::v2ConsoleBool::o: return #o;break;
		SWITCHCASE(unknown);
		SWITCHCASE(set);
		SWITCHCASE(unset);
		#undef SWITCHCASE
		default: return "unknown";
	}
}

char *ShowCmdToString(int iShowCmd)
{
	switch (iShowCmd)
	{
		#define SWITCHCASE(o) case (o): return #o;break;
		SWITCHCASE(SW_HIDE);
		SWITCHCASE(SW_SHOWNORMAL);
		//SWITCHCASE(SW_NORMAL);
		SWITCHCASE(SW_SHOWMINIMIZED);
		SWITCHCASE(SW_SHOWMAXIMIZED);
		//SWITCHCASE(SW_MAXIMIZE);
		SWITCHCASE(SW_SHOWNOACTIVATE);
		SWITCHCASE(SW_SHOW);
		SWITCHCASE(SW_MINIMIZE);
		SWITCHCASE(SW_SHOWMINNOACTIVE);
		SWITCHCASE(SW_SHOWNA);
		SWITCHCASE(SW_RESTORE);
		SWITCHCASE(SW_SHOWDEFAULT);
		SWITCHCASE(SW_FORCEMINIMIZE);
		//SWITCHCASE(SW_MAX);
		#undef SWITCHCASE
		default: return "unknown";
	}
}


//=====================================================================================================================================================================================================
// Main function
//=====================================================================================================================================================================================================
int __cdecl _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr;
	NT_CONSOLE_PROPS props = {0};
	TCHAR szFontFaceName[MAX_PATH] = _T("Terminal");

	//argv = CommandLineToArgvW(GetCommandLine(), &argc);

#ifdef _DEBUG
	printf("%3d\n", argc);
	for (int i=0 ; i<argc ; i++)
	{
		printf("%3d: \"%S\"\n", i, argv[i]);
	}
#endif

	memcpy(&props, &defaultProps, sizeof(props));

	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	bool showlogo = true;
	if (argc > 2)
	{
		for (int i=1 ; i<argc ; ++i)
		{
			if (0 == _tcscmp(argv[i], _T("-nl")))
			{
				showlogo = false;
			}
		}
	}

	if (showlogo)
	{
		printf("Shortcut " THIS_VERSION " Copyright(c) 2013-2023 SpongySoft Software " FLAVOR "(%zd-bit).\n", 8*sizeof(void*));
	}


	if (argc > 2)
	{
		_TCHAR szTargetFile[MAX_STRING] = {0};
		_TCHAR szWorkingFolder[MAX_STRING] = {0};
		_TCHAR szDescription[MAX_STRING] = {0};
		_TCHAR szArguments[MAX_STRING] = {0};
		_TCHAR szIcon[MAX_STRING] = {0};
		DWORD dwIcon = 0;
		ShortCut *pShortCut = NULL;

		if (0 == _tcscmp(argv[1], _T("create")))
		{
			pShortCut = ShortCut::Create(argv[2]);
		}
		else if (0 == _tcscmp(argv[1], _T("edit")))
		{
			pShortCut = ShortCut::Open(argv[2]);
		}
		else if (0 == _tcscmp(argv[1], _T("query")))
		{
			pShortCut = ShortCut::Open(argv[2]);

			if (!pShortCut)
			{
				fprintf(stderr, "Unable to open shortcut \"%S\"", argv[2]);
				return -1;
			}

			bool bProps = false;
			int iShowCmd = 0;

			if (FAILED(hr=pShortCut->GetTarget(szTargetFile, sizeof(szTargetFile)))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			if (FAILED(hr=pShortCut->GetArguments(szArguments, sizeof(szArguments)))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			if (FAILED(hr=pShortCut->GetDescription(szDescription, sizeof(szDescription)))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			if (FAILED(hr=pShortCut->GetWorkingDirectory(szWorkingFolder, sizeof(szWorkingFolder)))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			if (FAILED(hr=pShortCut->GetIcon(szIcon, sizeof(szIcon), (INT *)&dwIcon))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			if (FAILED(hr=pShortCut->GetShowCmd(iShowCmd))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			if (SUCCEEDED(hr=pShortCut->GetConsoleProps(&props, sizeof(props)))) { bProps = true; }

			printf("Shortcut file:  \"%S\"\n", argv[2]);
			printf("Target file:    \"%S\"\n", szTargetFile);
			printf("Working Folder: \"%S\"\n", szWorkingFolder);
			printf("Arguments:      \"%S\"\n", szArguments);
			printf("Description:    \"%S\"\n", szDescription);
			printf("Icon:           \"%S\" (%d)\n", szIcon, dwIcon);
			printf("Show Command:   \"%s\"\n", ShowCmdToString(iShowCmd));

			if (bProps)
			{
				printf("Props.wFillAttribute:          %d %d\n", props.wFillAttribute>>4,props.wFillAttribute&0xF);
				printf("Props.wPopupFillAttribute:     %d %d\n", props.wPopupFillAttribute>>4,props.wPopupFillAttribute&0xF);
				printf("Props.dwScreenBufferSize:      %d,%d\n", props.dwScreenBufferSize.X, props.dwScreenBufferSize.Y);
				printf("Props.dwWindowSize:            %d,%d\n", props.dwWindowSize.X, props.dwWindowSize.Y);
				printf("Props.dwWindowOrigin:          %d,%d\n", props.dwWindowOrigin.X, props.dwWindowOrigin.Y);
				printf("Props.nFont:                   %d\n", props.nFont);
				printf("Props.nInputBufferSize:        %d\n", props.nInputBufferSize);
				printf("Props.dwFontSize:              %d,%d\n", props.dwFontSize.X, props.dwFontSize.Y);
				printf("Props.uFontFamily:             %d\n", props.uFontFamily);
				printf("Props.uFontWeight:             %d\n", props.uFontWeight);
				printf("Props.FaceName:                %S\n", props.FaceName);
				printf("Props.uCursorSize:             %d\n", props.uCursorSize);
				printf("Props.bFullScreen:             %d\n", props.bFullScreen);
				printf("Props.bQuickEdit:              %d\n", props.bQuickEdit);
				printf("Props.bInsertMode:             %d\n", props.bInsertMode);
				printf("Props.bAutoPosition:           %d\n", props.bAutoPosition);
				printf("Props.uHistoryBufferSize:      %d\n", props.uHistoryBufferSize);
				printf("Props.uNumberOfHistoryBuffers: %d\n", props.uNumberOfHistoryBuffers);
				printf("Props.bHistoryNoDup:           %d\n", props.bHistoryNoDup);
				printf("Props.ColorTable[ 0]:          %08X\n", props.ColorTable[ 0]);
				printf("Props.ColorTable[ 1]:          %08X\n", props.ColorTable[ 1]);
				printf("Props.ColorTable[ 2]:          %08X\n", props.ColorTable[ 2]);
				printf("Props.ColorTable[ 3]:          %08X\n", props.ColorTable[ 3]);
				printf("Props.ColorTable[ 4]:          %08X\n", props.ColorTable[ 4]);
				printf("Props.ColorTable[ 5]:          %08X\n", props.ColorTable[ 5]);
				printf("Props.ColorTable[ 6]:          %08X\n", props.ColorTable[ 6]);
				printf("Props.ColorTable[ 7]:          %08X\n", props.ColorTable[ 7]);
				printf("Props.ColorTable[ 8]:          %08X\n", props.ColorTable[ 8]);
				printf("Props.ColorTable[ 9]:          %08X\n", props.ColorTable[ 9]);
				printf("Props.ColorTable[10]:          %08X\n", props.ColorTable[10]);
				printf("Props.ColorTable[11]:          %08X\n", props.ColorTable[11]);
				printf("Props.ColorTable[12]:          %08X\n", props.ColorTable[12]);
				printf("Props.ColorTable[13]:          %08X\n", props.ColorTable[13]);
				printf("Props.ColorTable[14]:          %08X\n", props.ColorTable[14]);
				printf("Props.ColorTable[15]:          %08X\n", props.ColorTable[15]);
			}

			for (int i = 0; i<static_cast<int>(ShortCut::v2ConsoleOption::maxvalue); ++i)
			{
				ShortCut::v2ConsoleBool b;

				if (FAILED(hr=pShortCut->GetV2ConsoleOption(static_cast<ShortCut::v2ConsoleOption>(i), &b))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				else
				{
					printf("%-21s          %s\n", ConsoleOptionToString(static_cast<ShortCut::v2ConsoleOption>(i)), ConsoleBoolToString(b));
				}
			}

			BOOL bAdmin;
			if (FAILED(hr=pShortCut->GetRunAsAdmin(bAdmin))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			printf("Run as:                        %s\n", bAdmin ? "Administrator" : "Standard User");

			// flags
			//DWORD dwFlags = pShortCut->GetFlags();
			//printf("Flags:                         %08X\n", dwFlags);
			//#define FLAG_PRINT(flag) if (dwFlags & flag) { printf("                               %s (0x%08X)\n", #flag, flag); dwFlags &= ~flag; }
			//
			//FLAG_PRINT(SLDF_HAS_ID_LIST);
			//FLAG_PRINT(SLDF_HAS_LINK_INFO);
			//FLAG_PRINT(SLDF_HAS_NAME);
			//FLAG_PRINT(SLDF_HAS_RELPATH);
			//FLAG_PRINT(SLDF_HAS_WORKINGDIR);
			//FLAG_PRINT(SLDF_HAS_ARGS);
			//FLAG_PRINT(SLDF_HAS_ICONLOCATION);
			//FLAG_PRINT(SLDF_UNICODE);
			//FLAG_PRINT(SLDF_FORCE_NO_LINKINFO);
			//FLAG_PRINT(SLDF_HAS_EXP_SZ);
			//FLAG_PRINT(SLDF_RUN_IN_SEPARATE);
			//FLAG_PRINT(SLDF_HAS_LOGO3ID);
			//FLAG_PRINT(SLDF_HAS_DARWINID);
			//FLAG_PRINT(SLDF_RUNAS_USER);
			//FLAG_PRINT(SLDF_HAS_EXP_ICON_SZ);
			//FLAG_PRINT(SLDF_NO_PIDL_ALIAS);
			//FLAG_PRINT(SLDF_FORCE_UNCNAME);
			//FLAG_PRINT(SLDF_RUN_WITH_SHIMLAYER);
			//FLAG_PRINT(SLDF_FORCE_NO_LINKTRACK);
			//FLAG_PRINT(SLDF_ENABLE_TARGET_METADATA);
			//FLAG_PRINT(SLDF_DISABLE_KNOWNFOLDER_RELATIVE_TRACKING);
			//
			//printf("Flags:                         %08X\n", dwFlags);


			pShortCut->Release();
			return 0;
		}
		else
		{
			usage();
		}

		if (!pShortCut)
		{
			fprintf(stderr, "Unable to open shortcut \"%S\"", argv[2]);
			return -1;
		}

		if (FAILED(hr=pShortCut->GetConsoleProps(&props, sizeof(props))))
		{
			memcpy(&props, &defaultProps, sizeof(props));
		}

		for (int i=3 ; i<argc ; ++i)
		{
			if (0 == _tcscmp(argv[i], _T("-t")))
			{
				if (++i < argc)
				{
					_tcscpy_s(szTargetFile, ARRAY_SIZE(szTargetFile), argv[i]);

					if (FAILED(hr=pShortCut->SetTarget(szTargetFile))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-w")))
			{
				if (++i < argc)
				{
					_tcscpy_s(szWorkingFolder, ARRAY_SIZE(szWorkingFolder), argv[i]);

					if (FAILED(hr=pShortCut->SetWorkingDirectory(szWorkingFolder))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-d")))
			{
				if (++i < argc)
				{
					_tcscpy_s(szDescription, ARRAY_SIZE(szDescription), argv[i]);

					if (FAILED(hr=pShortCut->SetDescription(szDescription))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-a")))
			{
				if (++i < argc)
				{
					_tcscpy_s(szArguments, ARRAY_SIZE(szArguments), argv[i]);

					while (++i < argc)
					{
						_tcscat_s(szArguments, ARRAY_SIZE(szArguments), _T(" "));
						_tcscat_s(szArguments, ARRAY_SIZE(szArguments), argv[i]);
					}

					if (FAILED(hr=pShortCut->SetArguments(szArguments))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-i")))
			{
				if (i+2 < argc)
				{
					_tcscpy_s(szIcon, ARRAY_SIZE(szIcon), argv[++i]);
					dwIcon = (DWORD)_ttoi(argv[++i]);

					if (FAILED(hr=pShortCut->SetIcon(szIcon, dwIcon))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-c")))
			{
				if (i+2 < argc)
				{
					DWORD index = (DWORD)_ttoi(argv[++i]);
					DWORD color = _htoi(argv[++i]);

					if (index < 16)
					{
						props.ColorTable[index] = color;
					}

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-ws")))
			{
				if (i+2 < argc)
				{
					props.dwWindowSize.X = (SHORT)_ttoi(argv[++i]);
					props.dwWindowSize.Y = (SHORT)_ttoi(argv[++i]);

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-bs")))
			{
				if (i+2 < argc)
				{
					props.dwScreenBufferSize.X = (SHORT)_ttoi(argv[++i]);
					props.dwScreenBufferSize.Y = (SHORT)_ttoi(argv[++i]);

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-qe")))
			{
				if (++i < argc)
				{
					props.bQuickEdit = (DWORD)_ttoi(argv[i]);
					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-im")))
			{
				if (++i < argc)
				{
					props.bQuickEdit = (DWORD)_ttoi(argv[i]);
					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-ra")))
			{
				if (++i < argc)
				{
					DWORD run_as_admin = (DWORD)_ttoi(argv[i]);
					if (FAILED(hr=pShortCut->SetRunAsAdmin(!(0==run_as_admin)))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-lw")))
			{
				if (++i < argc)
				{
					ShortCut::v2ConsoleOption option = ShortCut::v2ConsoleOption::lineselection;
					ShortCut::v2ConsoleBool value = ((DWORD)_ttoi(argv[i]) == 1) ? ShortCut::v2ConsoleBool::set : ShortCut::v2ConsoleBool::unset;

					if (FAILED(hr=pShortCut->SetV2ConsoleOption(option, value))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-v2")))
			{
				if (++i < argc)
				{
					ShortCut::v2ConsoleOption option = ShortCut::v2ConsoleOption::forcev2;
					ShortCut::v2ConsoleBool value = ((DWORD)_ttoi(argv[i]) == 1) ? ShortCut::v2ConsoleBool::set : ShortCut::v2ConsoleBool::unset;

					if (FAILED(hr=pShortCut->SetV2ConsoleOption(option, value))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }

					if (value == ShortCut::v2ConsoleBool::set)
					{
						if (FAILED(hr=pShortCut->SetV2ConsoleOption(ShortCut::v2ConsoleOption::filteronpaste, ShortCut::v2ConsoleBool::set))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
						if (FAILED(hr=pShortCut->SetV2ConsoleOption(ShortCut::v2ConsoleOption::windowtransparency, ShortCut::v2ConsoleBool::set))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
					}
					else
					{
						if (FAILED(hr=pShortCut->SetV2ConsoleOption(ShortCut::v2ConsoleOption::filteronpaste, ShortCut::v2ConsoleBool::unset))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
						if (FAILED(hr=pShortCut->SetV2ConsoleOption(ShortCut::v2ConsoleOption::windowtransparency, ShortCut::v2ConsoleBool::unset))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
					}
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-fn"))) // font face name
			{
				if (++i < argc)
				{
					_tcscpy_s(props.FaceName, ARRAY_SIZE(props.FaceName), argv[i]);

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-ff"))) // font family
			{
				if (++i < argc)
				{
					props.uFontFamily = (DWORD)_ttoi(argv[i]);

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-fw"))) // font weight
			{
				if (++i < argc)
				{
					props.uFontWeight = (DWORD)_ttoi(argv[i]);

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-fs"))) // font size
			{
				if (i+2 < argc)
				{
					props.dwFontSize.X = (short)_ttoi(argv[++i]);
					props.dwFontSize.Y = (short)_ttoi(argv[++i]);

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-fa"))) // fill attribute
			{
				if (i+2 < argc)
				{
					WORD fa1 = ((WORD)_ttoi(argv[++i]))&0xF;
					WORD fa2 = ((WORD)_ttoi(argv[++i]))&0xF;
					props.wFillAttribute = (fa1 << 4) | fa2;

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-pf"))) // popup fill attribute
			{
				if (i+2 < argc)
				{
					WORD fa1 = ((WORD)_ttoi(argv[++i]))&0xF;
					WORD fa2 = ((WORD)_ttoi(argv[++i]))&0xF;
					props.wPopupFillAttribute = (fa1 << 4) | fa2;

					if (FAILED(hr=pShortCut->SetConsoleProps(&props))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
				}
			}
			else if (0 == _tcscmp(argv[i], _T("-min"))) // show minimized
			{
				if (FAILED(hr=pShortCut->SetShowCmd(SW_SHOWMINNOACTIVE))) { fprintf(stderr, "Error 0x%08X on line %d\n", hr, __LINE__); }
			}
			else if (0 == _tcscmp(argv[i], _T("-nl")))
			{
			}
			else
			{
				// unknown argument...
				printf("Ignoring unknown option \"%S\"\n", argv[i]);
			}
		}

		if (FAILED(hr=pShortCut->Save())) { fprintf(stderr, "Error 0x%08X saving \"%S\": %s(%d)\n", hr, pShortCut->ShortcutFile(), __FILE__, __LINE__); }
		pShortCut->Release();
	}
	else
	{
		usage();
	}

	CoUninitialize();
	return 0;
}

