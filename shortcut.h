//=====================================================================================================================================================================================================
//
// SpongySoft ShortCut
// Copyright (C) SpongySoft. All rights reserved.
//
//=====================================================================================================================================================================================================
#define RELEASE(p) if (p) { (p)->Release(); (p)=NULL; }

namespace SpongySoft
{
	class ShortCut
	{
	public:
		enum class v2ConsoleBool : unsigned int
		{
			unknown,
			set,
			unset,
		};

		enum class v2ConsoleOption :unsigned int
		{
			forcev2,
			wraptext,
			filteronpaste,
			ctrlkeysdisabled,
			lineselection,
			windowtransparency,
			trimzeros,

			maxvalue,
		};

	public:
		static ShortCut* Create(LPCTSTR pszShortcutFile, LPCTSTR pszShortcutTarget, LPCTSTR pszShortcutArguments, LPCTSTR pszShortcutDescription, LPCTSTR pszShortcutWorkingDirectory);
		static ShortCut* Create(LPCTSTR pszShortcutFile);
		static ShortCut* Open(LPCTSTR pszShortcutFile);

	public:
		virtual HRESULT SetTarget(LPCTSTR pszShortcutTarget) = 0;
		virtual HRESULT SetArguments(LPCTSTR pszShortcutArguments) = 0;
		virtual HRESULT SetDescription(LPCTSTR pszShortcutDescription) = 0;
		virtual HRESULT SetWorkingDirectory(LPCTSTR pszShortcutWorkingDirectory) = 0;
		virtual HRESULT SetIcon(LPCTSTR pszShortcutIconFile, INT iIconIndex) = 0;
		virtual HRESULT SetConsoleProps(NT_CONSOLE_PROPS const *pConsoleProps) = 0;
		virtual HRESULT SetV2ConsoleOption(v2ConsoleOption option, v2ConsoleBool result) = 0;
		virtual HRESULT SetShowCmd(int iShowCmd) = 0;

		virtual HRESULT GetTarget(LPTSTR pszShortcutTarget, DWORD dwSizeInBytes) const = 0;
		virtual HRESULT GetArguments(LPTSTR pszShortcutArguments, DWORD dwSizeInBytes) const = 0;
		virtual HRESULT GetDescription(LPTSTR pszShortcutDescription, DWORD dwSizeInBytes) const = 0;
		virtual HRESULT GetWorkingDirectory(LPTSTR pszShortcutWorkingDirectory, DWORD dwSizeInBytes) const = 0;
		virtual HRESULT GetIcon(LPTSTR pszShortcutIconFile, DWORD dwSizeInBytes, INT *piIconIndex) const = 0;
		virtual HRESULT GetConsoleProps(NT_CONSOLE_PROPS *pConsoleProps, DWORD dwSize) const = 0;
		virtual HRESULT GetV2ConsoleOption(v2ConsoleOption option, v2ConsoleBool *result) const = 0;
		virtual HRESULT GetShowCmd(int &iShowCmd) const = 0;

		virtual HRESULT GetRunAsAdmin(BOOL &bRunAsAdmin) const = 0;
		virtual HRESULT SetRunAsAdmin(BOOL bRunAsAdmin) = 0;

//#ifdef _DEBUG
		virtual LPCTSTR Target(void) const = 0;
		virtual LPCTSTR Arguments(void) const = 0;
		virtual LPCTSTR Description(void) const = 0;
		virtual LPCTSTR WorkingDirectory(void) const = 0;
		virtual LPCTSTR IconFile(void) const = 0;
		virtual LPCTSTR ShortcutFile(void) const = 0;
		virtual INT IconIndex(void) const = 0;
		virtual const NT_CONSOLE_PROPS *ConsoleProps(void) const = 0;
//#endif

	public:
		virtual HRESULT Save(void) const = 0;

	public:
		virtual void Release(void) = 0;
	};
}

