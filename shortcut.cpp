#include "stdafx.h"
#include "shortcut.h"

#define MAX_STRING 4096
#define USE_CCOMPTR

#define ARRAY_SIZE(a) (sizeof(a)/sizeof((a)[0]))
#define copystring(dst,src) _tcscpy_s(dst, ARRAY_SIZE(dst), src)

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
#define PID_CONSOLE_FORCEV2            1
#define PID_CONSOLE_WRAPTEXT           2
#define PID_CONSOLE_FILTERONPASTE      3
#define PID_CONSOLE_CTRLKEYSDISABLED   4
#define PID_CONSOLE_LINESELECTION      5
#define PID_CONSOLE_WINDOWTRANSPARENCY 6

#define DEFINE_PROPERTYKEY_FORCE(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8, pid) extern "C" const PROPERTYKEY __declspec(selectany) name = { { l, w1, w2, { b1, b2,  b3,  b4,  b5,  b6,  b7,  b8 } }, pid }

DEFINE_PROPERTYKEY_FORCE(PKEY_Console_ForceV2,					0x0C570607, 0x0396, 0x43DE, 0x9D, 0x61, 0xE3, 0x21, 0xD7, 0xDF, 0x50, 0x26,	PID_CONSOLE_FORCEV2);
DEFINE_PROPERTYKEY_FORCE(PKEY_Console_WrapText,					0x0C570607, 0x0396, 0x43DE, 0x9D, 0x61, 0xE3, 0x21, 0xD7, 0xDF, 0x50, 0x26,	PID_CONSOLE_WRAPTEXT);
DEFINE_PROPERTYKEY_FORCE(PKEY_Console_FilterOnPaste,			0x0C570607, 0x0396, 0x43DE, 0x9D, 0x61, 0xE3, 0x21, 0xD7, 0xDF, 0x50, 0x26,	PID_CONSOLE_FILTERONPASTE);
DEFINE_PROPERTYKEY_FORCE(PKEY_Console_CtrlKeyShortcutsDisabled,	0x0C570607, 0x0396, 0x43DE, 0x9D, 0x61, 0xE3, 0x21, 0xD7, 0xDF, 0x50, 0x26,	PID_CONSOLE_CTRLKEYSDISABLED);
DEFINE_PROPERTYKEY_FORCE(PKEY_Console_LineSelection,			0x0C570607, 0x0396, 0x43DE, 0x9D, 0x61, 0xE3, 0x21, 0xD7, 0xDF, 0x50, 0x26,	PID_CONSOLE_LINESELECTION);
DEFINE_PROPERTYKEY_FORCE(PKEY_Console_WindowTransparency,		0x0C570607, 0x0396, 0x43DE, 0x9D, 0x61, 0xE3, 0x21, 0xD7, 0xDF, 0x50, 0x26,	PID_CONSOLE_WINDOWTRANSPARENCY);

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
#ifdef _DEBUG
#define CHK(statement) { hr = statement; if (FAILED(hr)) { printf("%s (%d): error 0x%08X\n", __FILE__, __LINE__, hr); return hr; } }
#else
#define CHK(statement) { hr = statement; if (FAILED(hr)) { return hr; } }
#endif

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
#pragma comment(lib, "propsys")

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
namespace SpongySoft
{
	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	static struct
	{
		const PROPERTYKEY *			propKey;
		ShortCut::v2ConsoleOption	option;
	} v2ConsoleOptionTable[] = {
		{ &PKEY_Console_ForceV2,					ShortCut::v2ConsoleOption::forcev2 },
		{ &PKEY_Console_WrapText,					ShortCut::v2ConsoleOption::wraptext },
		{ &PKEY_Console_FilterOnPaste,				ShortCut::v2ConsoleOption::filteronpaste },
		{ &PKEY_Console_CtrlKeyShortcutsDisabled,	ShortCut::v2ConsoleOption::ctrlkeysdisabled },
		{ &PKEY_Console_LineSelection,				ShortCut::v2ConsoleOption::lineselection },
		{ &PKEY_Console_WindowTransparency,			ShortCut::v2ConsoleOption::windowtransparency },
	};

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	class _ShortCut : public ShortCut
	{
		virtual HRESULT SetTarget(LPCTSTR pszShortcutTarget);
		virtual HRESULT SetArguments(LPCTSTR pszShortcutArguments);
		virtual HRESULT SetDescription(LPCTSTR pszShortcutDescription);
		virtual HRESULT SetWorkingDirectory(LPCTSTR pszShortcutWorkingDirectory);
		virtual HRESULT SetIcon(LPCTSTR pszShortcutIconFile, INT iIconIndex);
		virtual HRESULT SetConsoleProps(NT_CONSOLE_PROPS const *pConsoleProps);
		virtual HRESULT SetV2ConsoleOption(v2ConsoleOption option, v2ConsoleBool result);

		virtual HRESULT GetTarget(LPTSTR pszShortcutTarget, DWORD dwSizeInBytes) const;
		virtual HRESULT GetArguments(LPTSTR pszShortcutArguments, DWORD dwSizeInBytes) const;
		virtual HRESULT GetDescription(LPTSTR pszShortcutDescription, DWORD dwSizeInBytes) const;
		virtual HRESULT GetWorkingDirectory(LPTSTR pszShortcutWorkingDirectory, DWORD dwSizeInBytes) const;
		virtual HRESULT GetIcon(LPTSTR pszShortcutIconFile, DWORD dwSizeInBytes, INT *piIconIndex) const;
		virtual HRESULT GetConsoleProps(NT_CONSOLE_PROPS *pConsoleProps, DWORD dwSize) const;
		virtual HRESULT GetV2ConsoleOption(v2ConsoleOption option, v2ConsoleBool *result) const;

		virtual DWORD GetFlags(void) const;
		virtual void SetFlags(DWORD flags);
		virtual void SetFlags(DWORD flagsOr, DWORD flagsAnd);
		virtual HRESULT GetRunAsAdmin(BOOL &bRunAsAdmin) const;
		virtual HRESULT SetRunAsAdmin(BOOL bRunAsAdmin);

#ifdef _DEBUG
		virtual LPCTSTR Target(void) const { return this->szShortcutTarget; }
		virtual LPCTSTR Arguments(void) const { return this->szShortcutArguments; }
		virtual LPCTSTR Description(void) const { return this->szShortcutDescription; }
		virtual LPCTSTR WorkingDirectory(void) const { return this->szShortcutWorkingDirectory; }
		virtual LPCTSTR IconFile(void) const { return this->szShortcutIconFile; }
		virtual INT IconIndex(void) const  { return this->iIcon; }
		virtual const NT_CONSOLE_PROPS *ConsoleProps(void) const { return &this->props; }
#endif

	public:
		virtual HRESULT Save(void) const;

	public:
		virtual void Release(void);

	private:
		_ShortCut(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory);
		_ShortCut(LPCTSTR pszShortcutFile);
		virtual ~_ShortCut();

	private:
		void _internal_init(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory);
		HRESULT Load();

	private:
		TCHAR   			szShortcutFile[MAX_STRING];
		TCHAR   			szShortcutTarget[MAX_STRING];
		TCHAR   			szShortcutArguments[MAX_STRING];
		TCHAR   			szShortcutDescription[MAX_STRING];
		TCHAR   			szShortcutWorkingDirectory[MAX_STRING];
		TCHAR   			szShortcutIconFile[MAX_STRING];
		int					iIcon;
		NT_CONSOLE_PROPS	props;
		bool				propsAreValid;
		DWORD				dwFlags;
		bool				runAsAdmin;
		v2ConsoleBool		v2ConsoleOptions[v2ConsoleOption::maxvalue];

	friend ShortCut* ShortCut::Create(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory);
	friend ShortCut* ShortCut::Create(LPCTSTR pszShortcutFile);
	friend ShortCut* ShortCut::Open(LPCTSTR pszShortcutFile);
	};

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	ShortCut* ShortCut::Create(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory)
	{
		_ShortCut *pShortCut = new _ShortCut(pszShortcutFile, pszShortcutTarget, pszShortcutArguments, pszShortcutDescription, pszShortcutWorkingDirectory);

		if (SUCCEEDED(pShortCut->Load()))
		{
			pShortCut->Release();
			return NULL;
		}

		return pShortCut;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	ShortCut* ShortCut::Create(LPCTSTR pszShortcutFile)
	{
		_ShortCut *pShortCut = new _ShortCut(pszShortcutFile);

		if (SUCCEEDED(pShortCut->Load()))
		{
			pShortCut->Release();
			return NULL;
		}

		return pShortCut;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	ShortCut* ShortCut::Open(LPCTSTR pszShortcutFile)
	{
		_ShortCut *pShortCut = new _ShortCut(pszShortcutFile);

		if (SUCCEEDED(pShortCut->Load()))
		{
			return pShortCut;
		}
		else
		{
			pShortCut->Release();
		}

		return NULL;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	void _ShortCut::_internal_init(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory)
	{
		copystring(this->szShortcutFile,				pszShortcutFile);
		copystring(this->szShortcutTarget,				pszShortcutTarget);
		copystring(this->szShortcutArguments,			pszShortcutArguments);
		copystring(this->szShortcutDescription,			pszShortcutDescription);
		copystring(this->szShortcutWorkingDirectory,	pszShortcutWorkingDirectory);
		copystring(this->szShortcutIconFile,			_T(""));
		this->iIcon = 0;
		this->propsAreValid = false;
		this->dwFlags = 0;

		memset(&this->props, 0, sizeof(this->props));
		memset(&this->v2ConsoleOptions, 0, sizeof(this->v2ConsoleOptions));
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	_ShortCut::_ShortCut(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory)
	{
		this->_internal_init(pszShortcutFile, pszShortcutTarget, pszShortcutArguments, pszShortcutDescription, pszShortcutWorkingDirectory);
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	_ShortCut::_ShortCut(LPCTSTR pszShortcutFile)
	{
		this->_internal_init(pszShortcutFile, _T(""), _T(""), _T(""), _T(""));
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	_ShortCut::~_ShortCut()
	{
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	void _ShortCut::Release(void)
	{
		delete this;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetTarget(LPCTSTR pszShortcutTarget)
	{
		copystring(this->szShortcutTarget,				pszShortcutTarget);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetArguments(LPCTSTR pszShortcutArguments)
	{
		copystring(this->szShortcutArguments,			pszShortcutArguments);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetDescription(LPCTSTR pszShortcutDescription)
	{
		copystring(this->szShortcutDescription,			pszShortcutDescription);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetWorkingDirectory(LPCTSTR pszShortcutWorkingDirectory)
	{
		copystring(this->szShortcutWorkingDirectory,	pszShortcutWorkingDirectory);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetIcon(LPCTSTR pszShortcutIconFile, INT iIconIndex)
	{
		copystring(this->szShortcutIconFile,			pszShortcutIconFile);
		this->iIcon = iIconIndex;
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetTarget(LPTSTR pszShortcutTarget, DWORD dwSizeInBytes) const
	{
		_tcscpy_s(pszShortcutTarget, dwSizeInBytes/sizeof(TCHAR), this->szShortcutTarget);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetArguments(LPTSTR pszShortcutArguments, DWORD dwSizeInBytes) const
	{
		_tcscpy_s(pszShortcutArguments, dwSizeInBytes/sizeof(TCHAR), this->szShortcutArguments);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetDescription(LPTSTR pszShortcutDescription, DWORD dwSizeInBytes) const
	{
		_tcscpy_s(pszShortcutDescription, dwSizeInBytes/sizeof(TCHAR), this->szShortcutDescription);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetWorkingDirectory(LPTSTR pszShortcutWorkingDirectory, DWORD dwSizeInBytes) const
	{
		_tcscpy_s(pszShortcutWorkingDirectory, dwSizeInBytes/sizeof(TCHAR), this->szShortcutWorkingDirectory);
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetIcon(LPTSTR pszShortcutFile, DWORD dwSizeInBytes, INT *piIconIndex) const
	{
		_tcscpy_s(pszShortcutFile, dwSizeInBytes/sizeof(TCHAR), this->szShortcutIconFile);
		if (NULL != piIconIndex)
		{
			*piIconIndex = this->iIcon;
		}

		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetConsoleProps(NT_CONSOLE_PROPS *pConsoleProps, DWORD dwSize) const
	{
		if (this->propsAreValid)
		{
			if (dwSize < sizeof(this->props))
			{
				return E_OUTOFMEMORY;
			}

			if (NULL == pConsoleProps)
			{
				return E_POINTER;
			}

			memcpy(pConsoleProps, &this->props, sizeof(this->props));

			return S_OK;
		}
		else
		{
			return E_INVALIDARG;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetConsoleProps(NT_CONSOLE_PROPS const *pConsoleProps)
	{
		if (NULL == pConsoleProps)
		{
			return E_POINTER;
		}

		memcpy(&this->props, pConsoleProps, sizeof(this->props));
		this->propsAreValid = true;
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::Save(void) const
	{
		HRESULT hr = E_FAIL;
		#ifdef USE_CCOMPTR
		CComPtr<IShellLink> pShellLink = nullptr;
		#else
		IShellLink* pShellLink;
		#endif

		if (SUCCEEDED(hr=CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&pShellLink)))
		{
			#ifdef USE_CCOMPTR
			#else
			#endif

			#ifdef USE_CCOMPTR
			CComPtr<IPersistFile> pPersistFile = nullptr;
			#else
			IPersistFile* pPersistFile = nullptr;
			#endif

			if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_IPersistFile, (LPVOID *)&pPersistFile)))
			{
				pShellLink->SetPath(this->szShortcutTarget);
				pShellLink->SetArguments(this->szShortcutArguments);
				pShellLink->SetDescription(this->szShortcutDescription);
				pShellLink->SetWorkingDirectory(this->szShortcutWorkingDirectory);
				pShellLink->SetIconLocation(this->szShortcutIconFile, this->iIcon);

				// v2 console options
				if (true)
				{
					#ifdef USE_CCOMPTR
					CComPtr<IPropertyStore> pPropStoreLnk;
					//if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_PPV_ARGS(&pPropStoreLnk))))
					CHK(pShellLink->QueryInterface(IID_PPV_ARGS(&pPropStoreLnk)));
					#else
					IPropertyStore* pPropStoreLnk = nullptr;
					//if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_IPropertyStore, (LPVOID *)(&pPropStoreLnk))))
					CHK(pShellLink->QueryInterface(IID_PPV_ARGS(&pPropStoreLnk)));
					#endif
					{
						for (unsigned int i=0 ; i<static_cast<unsigned int>(v2ConsoleOption::maxvalue) ; ++i)
						{
							if (this->v2ConsoleOptions[i] != v2ConsoleBool::unknown)
							{
								bool bValue = (this->v2ConsoleOptions[i] == v2ConsoleBool::set);

								PROPVARIANT propvar;
								//if (SUCCEEDED(hr=InitPropVariantFromBoolean(bValue, &propvar)))
								CHK(InitPropVariantFromBoolean(bValue, &propvar));
								CHK(pPropStoreLnk->SetValue(*v2ConsoleOptionTable[i].propKey, propvar));
							}
						}
					}
				}

				// console properties
				if (true)
				{
					#ifdef USE_CCOMPTR
					CComPtr<IShellLinkDataList> pDataList = nullptr;
					#else
					IShellLinkDataList *pDataList = nullptr;
					#endif

					if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_IShellLinkDataList, (LPVOID *)&pDataList)))
					{
						if (this->propsAreValid)
						{
							if (SUCCEEDED(hr=pDataList->RemoveDataBlock(NT_CONSOLE_PROPS_SIG)))
							{
								if (SUCCEEDED(hr=pDataList->AddDataBlock(const_cast<NT_CONSOLE_PROPS *>(&this->props))))
								{
									int a=0;
								}
							}
						}

						DWORD dwFlags;
						if (SUCCEEDED(hr=pDataList->GetFlags(&dwFlags)))
						{
							if (this->runAsAdmin)
							{
#ifdef _DEBUG
								printf("%s (%d)\n", __FILE__,__LINE__);
#endif
								dwFlags |= SLDF_RUNAS_USER;
							}
							else
							{
#ifdef _DEBUG
								printf("%s (%d)\n", __FILE__,__LINE__);
#endif
								dwFlags &= ~SLDF_RUNAS_USER;
							}

							// other flags
							dwFlags |= SLDF_FORCE_NO_LINKINFO;
							dwFlags |= SLDF_FORCE_NO_LINKTRACK;
							dwFlags |= SLDF_DISABLE_LINK_PATH_TRACKING;
#ifdef _DEBUG
							printf("%s (%d)\n", __FILE__,__LINE__);
#endif
							if (SUCCEEDED(hr=pDataList->SetFlags(dwFlags)))
							{
#ifdef _DEBUG
								printf("%s (%d)\n", __FILE__,__LINE__);
#endif
							}

							if (SUCCEEDED(hr=pDataList->GetFlags(&dwFlags)))
							{
#ifdef _DEBUG
								printf("Flags: 0x%08X %s (%d)\n", dwFlags, __FILE__,__LINE__);
#endif
							}
						}


						#ifndef USE_CCOMPTR
						RELEASE(pDataList);
						#endif
					}
					else
					{
						// could not get the shell link data list object...
						int a=0;
					}
				}

				const WCHAR *pszShortcutFileName;
#ifdef UNICODE
				pszShortcutFileName = this->szShortcutFile;
#else
				WCHAR wszShortcutFile[MAX_STRING];
				MultiByteToWideChar(CP_ACP, 0, this->szShortcutFile, -1, wszShortcutFile, MAX_STRING);
				pszShortcutFileName = wszShortcutFile;
#endif
				if (SUCCEEDED(hr=pPersistFile->Save(pszShortcutFileName, 0)))
				{
				}

				#ifndef USE_CCOMPTR
				RELEASE(pPersistFile);
				#endif
			}

			#ifndef USE_CCOMPTR
			RELEASE(pShellLink);
			#endif
		}

		return hr;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::Load(void)
	{
		HRESULT hr = E_FAIL;
		#ifdef USE_CCOMPTR
		CComPtr<IShellLink> pShellLink = nullptr;
		#else
		IShellLink* pShellLink;
		#endif

		if (SUCCEEDED(hr=CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&pShellLink)))
		{
			#ifdef USE_CCOMPTR
			CComPtr<IPersistFile> pPersistFile = nullptr;
			#else
			IPersistFile* pPersistFile = nullptr;
			#endif

			#ifdef USE_CCOMPTR
			if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_PPV_ARGS(&pPersistFile))))
			#else
			if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_IPersistFile, (LPVOID *)&pPersistFile)))
			#endif
			{
				const WCHAR *pszShortcutFileName;
#ifdef UNICODE
				pszShortcutFileName = this->szShortcutFile;
#else
				WCHAR wszShortcutFile[MAX_STRING];
				MultiByteToWideChar(CP_ACP, 0, this->szShortcutFile, -1, wszShortcutFile, MAX_STRING);
				pszShortcutFileName = wszShortcutFile;
#endif
				if (SUCCEEDED(hr=pPersistFile->Load(pszShortcutFileName, 0)))
				{
					#ifdef USE_CCOMPTR
					CComPtr<IShellLinkDataList> pDataList = nullptr;
					#else
					IShellLinkDataList *pDataList = nullptr;
					#endif

					// v2 console options
					if (true)
					{
						#ifdef USE_CCOMPTR
						CComPtr<IPropertyStore> pPropStoreLnk;
						if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_PPV_ARGS(&pPropStoreLnk))))
						#else
						IPropertyStore* pPropStoreLnk = nullptr;
						if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_IPropertyStore, (LPVOID *)(&pPropStoreLnk))))
						#endif
						{
							for (unsigned int i=0 ; i<static_cast<unsigned int>(v2ConsoleOption::maxvalue) ; ++i)
							{
								PROPVARIANT propvar;
								BOOL bValue;

								CHK(pPropStoreLnk->GetValue(*v2ConsoleOptionTable[i].propKey, &propvar));
								CHK(PropVariantToBoolean(propvar, &bValue));
								this->v2ConsoleOptions[i] = bValue ? v2ConsoleBool::set : v2ConsoleBool::unset;
							}
						}
					}

					if (SUCCEEDED(hr=pShellLink->QueryInterface(IID_IShellLinkDataList, (LPVOID *)&pDataList)))
					{
						NT_CONSOLE_PROPS *pprops;

						if (SUCCEEDED(hr = pDataList->CopyDataBlock(NT_CONSOLE_PROPS_SIG, (void **)&pprops)))
						{
							if (NULL != pprops)
							{
								memcpy(&this->props, pprops, sizeof(this->props));
								LocalFree(pprops);
								this->propsAreValid = true;
							}
						}
						else
						{
							// there's no props...
							hr = S_OK;
							this->propsAreValid = false;
						}

						if (SUCCEEDED(hr=pDataList->GetFlags(&this->dwFlags)))
						{
						}
						else
						{
							this->dwFlags = 0xFFFFFFFF;
						}

						this->runAsAdmin = SLDF_RUNAS_USER == (this->dwFlags & SLDF_RUNAS_USER);
#ifdef _DEBUG
						printf("%s (%d) (%d)\n", __FILE__, __LINE__, this->runAsAdmin);
#endif

						#ifndef USE_CCOMPTR
						RELEASE(pDataList);
						#endif
					}
					else
					{
						// there's no props...
						this->propsAreValid = false;
					}

					pShellLink->GetArguments(this->szShortcutArguments,					ARRAY_SIZE(this->szShortcutArguments));
					pShellLink->GetDescription(this->szShortcutDescription,				ARRAY_SIZE(this->szShortcutDescription));
					pShellLink->GetIconLocation(this->szShortcutIconFile,				ARRAY_SIZE(this->szShortcutIconFile), &this->iIcon);
					pShellLink->GetPath(this->szShortcutTarget,							ARRAY_SIZE(this->szShortcutTarget), NULL, 0);
					pShellLink->GetWorkingDirectory(this->szShortcutWorkingDirectory,	ARRAY_SIZE(this->szShortcutWorkingDirectory));
				}

				#ifndef USE_CCOMPTR
				RELEASE(pPersistFile);
				#endif
			}

			#ifndef USE_CCOMPTR
			RELEASE(pShellLink);
			#endif
		}

		return hr;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	DWORD _ShortCut::GetFlags(void) const
	{
		return this->dwFlags;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	void _ShortCut::SetFlags(DWORD flags)
	{
		this->dwFlags = flags;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	void _ShortCut::SetFlags(DWORD flagsOr, DWORD flagsAnd)
	{
		this->dwFlags &= flagsAnd;
		this->dwFlags |= flagsOr;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetRunAsAdmin(BOOL &bRunAsAdmin) const
	{
		bRunAsAdmin = this->runAsAdmin ? TRUE : FALSE;
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetRunAsAdmin(BOOL bRunAsAdmin)
	{
		if (bRunAsAdmin)
		{
			this->runAsAdmin = true;
		}
		else
		{
			this->runAsAdmin = false;
		}
		return S_OK;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::SetV2ConsoleOption(v2ConsoleOption option, v2ConsoleBool value)
	{
		if (option <= v2ConsoleOption::maxvalue)
		{
			this->v2ConsoleOptions[static_cast<unsigned int>(option)] = value;
			return S_OK;
		}

		return E_INVALIDARG;;
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	HRESULT _ShortCut::GetV2ConsoleOption(v2ConsoleOption option, v2ConsoleBool *result) const
	{
		if (nullptr == result)
		{
			return E_INVALIDARG;
		}

		if (option <= v2ConsoleOption::maxvalue)
		{
			*result = this->v2ConsoleOptions[static_cast<unsigned int>(option)];
			return S_OK;
		}

		return E_INVALIDARG;;
	}

}

