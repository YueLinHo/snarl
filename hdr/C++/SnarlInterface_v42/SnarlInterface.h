﻿#ifndef SNARL_INTERFACE_V42
#define SNARL_INTERFACE_V42

#ifdef __MINGW32__
	#define MINGW_HAS_SECURE_API
#endif 

#include <tchar.h>
#include <windows.h>
#include <cstdio>
#include <vector>
#include <sstream>


#ifndef SMTO_NOTIMEOUTIFNOTHUNG
	#define SMTO_NOTIMEOUTIFNOTHUNG 8
#endif


namespace Snarl {
	namespace V42 {

	static const LPCTSTR SnarlWindowClass = _T("w>Snarl");
	static const LPCTSTR SnarlWindowTitle = _T("Snarl");

	static const LPCTSTR SnarlGlobalMsg = _T("SnarlGlobalEvent");
	static const LPCTSTR SnarlAppMsg = _T("SnarlAppMessage");

	static const DWORD WM_SNARLTEST = WM_USER + 237;

	// Enums put in own namespace, because ANSI C++ doesn't decorate enums with tagname :(
	namespace SnarlEnums {

		/// <summary>
		/// Global event identifiers.
		/// Identifiers marked with a '*' are sent by Snarl in two ways:
		///   1. As a broadcast message (uMsg = 'SNARL_GLOBAL_MSG')
		///   2. To the window registered in snRegisterConfig() or snRegisterConfig2()
		///      (uMsg = reply message specified at the time of registering)
		/// In both cases these values appear in wParam.
		///   
		/// Identifiers not marked are not broadcast; they are simply sent to the application's registered window.
		/// </summary>
		enum GlobalEvent
		{
			SnarlLaunched = 1,      // Snarl has just started running*
			SnarlQuit = 2,          // Snarl is about to stop running*
			SnarlStopped = 3,       // sent when stopped by user
			SnarlStarted = 4,       // sent when started by user

			// R2.4 Beta3
			SnarlUserAway,          // away mode was enabled
			SnarlUserBack,          // away mode was disabled
		};

		enum SnarlStatus
		{
			Success = 0,

			// Win32 callbacks (renamed under V42)
			CallbackRightClick = 32,           // Deprecated as of V42, ex. SNARL_NOTIFICATION_CLICKED/SNARL_NOTIFICATION_CANCELLED
			CallbackTimedOut,
			CallbackInvoked,               // left clicked and no default callback assigned
			CallbackMenuSelected,          // HIWORD(wParam) contains 1-based menu item index
			CallbackMiddleClick,                // Deprecated as of V42
			CallbackClosed,


			// critical errors
			ErrorFailed = 101,             // miscellaneous failure
			ErrorUnknownCommand,           // specified command not recognised
			ErrorTimedOut,                 // Snarl took too long to respond
			//104 gen critical #4
			//105 gen critical #5
			ErrorBadSocket = 106,          // invalid socket (or some other socket-related error)
			ErrorBadPacket = 107,          // badly formed request
			//108 net critical #3
			ErrorArgMissing = 109,         // required argument missing
			ErrorSystem,                   // internal system error
			//120 libsnarl critical block
			ErrorAccessDenied = 121,       // libsnarl only

			// warnings
			ErrorNotRunning = 201,         // Snarl handling window not found
			ErrorNotRegistered,
			ErrorAlreadyRegistered,        // not used yet; sn41RegisterApp() returns existing token
			ErrorClassAlreadyExists,       // not used yet
			ErrorClassBlocked,
			ErrorClassNotFound,
			ErrorNotificationNotFound,
			ErrorFlooding,                 // notification generated by same class within quantum
			ErrorDoNotDisturb,             // DnD mode is in effect was not logged as missed
			ErrorCouldNotDisplay,          // not enough space on-screen to display notification
			ErrorAuthFailure,              // password mismatch

			// informational
			WasMerged = 251,               // notification was merged, returned token is the one we merged with

			// callbacks
			NotifyGone = 301,              // reserved for future use

			// The following are currently specific to SNP 2.0 and are effectively the
			// Win32 SNARL_CALLBACK_nnn constants with 270 added to them

			// SNARL_NOTIFY_CLICK = 302    // indicates notification was right-clicked (deprecated as of V42)
			NotifyExpired = 303,
			NotifyInvoked = 304,           // note this was "ACK" in a previous life
			NotifyMenu,                    // indicates an item was selected from user-defined menu (deprecated as of V42)
			// SNARL_NOTIFY_EX_CLICK       // user clicked the middle mouse button (deprecated as of V42)
			// SNARL_NOTIFY_CLOSED         // user clicked the notification's close gadget

			// the following is generic to SNP and the Win32 API
			NotifyAction = 308,             // user picked an action from the list, the data value will indicate which one


			// C++ interface custom errors- not part of official API!
			ErrorCppInterface = 1001
		};

	} // namespace SnarlEnums

	// ----------------------------------------------------------------------------------------
	/// SnarlParameterList class definition - Helper class, not meant for broad use
	// ----------------------------------------------------------------------------------------
	template<class T>
	class SnarlParameterList
	{
	public:
		typedef std::pair<std::basic_string<T>, std::basic_string<T>> Type;

		SnarlParameterList()
		{
		}
		
		explicit SnarlParameterList(int initialCapacity)
		{
			list.reserve(initialCapacity);
		}

		void Add(const T* _key, const T* _value)
		{
			if (_value != NULL)
				list.push_back(std::pair<std::basic_string<T>, std::basic_string<T>>(std::basic_string<T>(_key), std::basic_string<T>(_value)));
		}

		void Add(const T* _key, LONG32 _value)
		{
			std::basic_stringstream<T> valStr;
			valStr << _value;
			list.push_back(std::pair<std::basic_string<T>, std::basic_string<T>>(std::basic_string<T>(_key), valStr.str()));
		}
		
		void Add(const T* _key, void* _value)
		{
			if (_value != NULL)
			{
				std::basic_stringstream<T> valStr;
				valStr << (INT_PTR)_value; // Uckly hack, to get stringstream to print void* as decimal not hex

				list.push_back(std::pair<std::basic_string<T>, std::basic_string<T>>(std::basic_string<T>(_key), valStr.str()));
			}
		}
		
		const std::vector<Type>& GetList() const
		{
			return list;
		}

	private:
		std::vector<std::pair<std::basic_string<T>, std::basic_string<T>>> list;
	};

	// ----------------------------------------------------------------------------------------
	// SnarlInterface class definition
	// ----------------------------------------------------------------------------------------
	class SnarlInterface
	{
	public:
		/// <summary>Requests strings known by Snarl</summary>
		class Requests
		{
		public:
			static LPCSTR  AddActionA()    { return  "addaction"; }
			static LPCWSTR AddActionW()    { return L"addaction"; }
			static LPCSTR  AddClassA()     { return  "addclass"; }
			static LPCWSTR AddClassW()     { return L"addclass"; }
			static LPCSTR  ClearActionsA() { return  "clearactions"; }
			static LPCWSTR ClearActionsW() { return L"clearactions"; }
			static LPCSTR  ClearClassesA() { return  "clearclasses"; }
			static LPCWSTR ClearClassesW() { return L"clearclasses"; }
			static LPCSTR  HelloA()        { return  "hello"; }
			static LPCWSTR HelloW()        { return L"hello"; }
			static LPCSTR  HideA()         { return  "hide"; }
			static LPCWSTR HideW()         { return L"hide"; }
			static LPCSTR  IsVisibleA()    { return  "isvisible"; }
			static LPCWSTR IsVisibleW()    { return L"isvisible"; }
			static LPCSTR  NotifyA()       { return  "notify"; }
			static LPCWSTR NotifyW()       { return L"notify"; }
			static LPCSTR  RegisterA()     { return  "reg"; } // register
			static LPCWSTR RegisterW()     { return L"reg"; }
			static LPCSTR  RemoveClassA()  { return  "remclass"; }
			static LPCWSTR RemoveClassW()  { return L"remclass"; }
			static LPCSTR  UnregisterA()   { return  "unregister"; }
			static LPCWSTR UnregisterW()   { return L"unregister"; }
			static LPCSTR  UpdateAppA()    { return  "updateapp"; }
			static LPCWSTR UpdateAppW()    { return L"updateapp"; }
			static LPCSTR  UpdateA()       { return  "update"; }
			static LPCWSTR UpdateW()       { return L"update"; }
			static LPCSTR  VersionA()      { return  "version"; }
			static LPCWSTR VersionW()      { return L"version"; }
		};

		SnarlInterface();
		virtual ~SnarlInterface();

		// ------------------------------------------------------------------------------------
		// Static functions
		// ------------------------------------------------------------------------------------

		// Use FreeString, when SnarlInterface returns a null terminated string pointer
		static LPTSTR AllocateString(size_t n) { return new TCHAR[n]; }
		static void FreeString(LPSTR str)      { delete [] str; str = NULL; }
		static void FreeString(LPCSTR str)     { delete [] str; }
		static void FreeString(LPWSTR str)     { delete [] str; str = NULL; }
		static void FreeString(LPCWSTR str)    { delete [] str; }

		/// <summary>Send message to Snarl.</summary>
		/// <param name='request'>The request string. If using unicode version, the string will be UTF8 encoded before sending.</param>
		/// <param name='replyTimeout'>Time to wait before timeout - Default = 1000</param>
		/// <returns>
		///   Return zero or positive on Success.
		///   Negative on failure. (Get error code by abs(return_value))
		///   <see>http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Windows_API#Return_Value</see>
		/// </returns>
		static LONG32 DoRequest(LPCSTR request, UINT replyTimeout = 1000);
		static LONG32 DoRequest(LPCWSTR request, UINT replyTimeout = 1000);

		/// <summary>Returns the global Snarl Application message  (V39)</summary>
		/// <returns>Returns Snarl application registered message.</returns>
		static UINT AppMsg();

		/// <summary>
		///     Returns the value of Snarl's global registered message.
		///     Notes:
		///       Snarl registers SNARL_GLOBAL_MSG during startup which it then uses to communicate
		///       with all running applications through a Windows broadcast message. This function can
		///       only fail if for some reason the Windows RegisterWindowMessage() function fails
		///       - given this, this function *cannnot* be used to test for the presence of Snarl.
		/// </summary>
		/// <returns>A 16-bit value (translated to 32-bit) which is the registered Windows message for Snarl.</returns>
		static UINT Broadcast();

		/// <summary>
		///     Get the path to where Snarl is installed.
		///     ** Remember to call <see cref="FreeString(LPSTR)" /> on the returned string !!!
		/// </summary>
		/// <returns>Returns the path to where Snarl is installed.</returns>
		/// <remarks>This is a V39 API method.</remarks>
		static LPCTSTR GetAppPath();

		/// <summary>
		///     Get the path to where the default Snarl icons are located.
		///     <para>** Remember to call <see cref="FreeString(LPSTR)" /> on the returned string !!!</para>
		/// </summary>
		/// <returns>Returns the path to where the default Snarl icons are located.</returns>
		/// <remarks>This is a V39 API method.</remarks>
		static LPCTSTR GetIconsPath();

		/// <summary>Returns a handle to the Snarl Dispatcher window  (V37)</summary>
		/// <returns>Returns handle to Snarl Dispatcher window, or zero if it's not found.</returns>
		/// <remarks>This is now the preferred way to test if Snarl is actually running.</remarks>
		static HWND GetSnarlWindow();

		/// <summary>Get Snarl version, if it is running.</summary>
		/// <returns>Returns a number indicating Snarl version.</returns>
		static LONG32 GetVersion();

		/// <summary>Check whether Snarl is running</summary>
		/// <returns>Returns true if Snarl system was found running.</returns>
		static BOOL IsSnarlRunning();

		
		// ------------------------------------------------------------------------------------


		/// <summary>Adds an action to an existing (on-screen or in the missed list) notification.</summary>
		LONG32 AddAction(LONG32 msgToken, LPCSTR  label, LPCSTR  cmd);
		LONG32 AddAction(LONG32 msgToken, LPCWSTR label, LPCWSTR cmd);
		
		/// <summary>Add a notification class to Snarl.</summary>
		LONG32 AddClass(LPCSTR classId, LPCSTR name, LPCSTR title = NULL, LPCSTR text = NULL, LPCSTR icon = NULL, LPCSTR sound = NULL, LONG32 duration = NULL, LPCSTR callback = NULL, bool enabled = true);
		LONG32 AddClass(LPCWSTR classId, LPCWSTR name, LPCWSTR title = NULL, LPCWSTR text = NULL, LPCWSTR icon = NULL, LPCWSTR sound = NULL, LONG32 duration = -1, LPCWSTR callback = NULL, bool enabled = true);

		/// <summary>Remove all notification classes in one call.</summary>
		LONG32 ClearActions(LONG32 msgToken);

		/// <summary>Remove all notification classes in one call.</summary>
		LONG32 ClearClasses();

		/// <summary>GetLastMsgToken() returns token of the last message sent to Snarl.</summary>
		/// <returns>Returns message token of last message.</returns>
		/// <remarks>This function is not in the official API!</remarks>
		LONG32 GetLastMsgToken() const;

		/// <summary>Hide a Snarl notification.</summary>
		LONG32 Hide(LONG32 msgToken);

		/// <summary>Test if a Snarl notification is visible.</summary>
		LONG32 IsVisible(LONG32 msgToken);

		/// <summary>Show a Snarl notification.</summary>
		/// <returns>Returns the notification token or negative on failure.</returns>
		/// <remarks>You can use <see cref="GetLastMsgToken()" /> to get the last token.</remarks>
		LONG32 Notify(LPCSTR classId = NULL, LPCSTR title = NULL, LPCSTR text = NULL, LONG32 timeout = -1, LPCSTR iconPath = NULL, LPCSTR iconBase64 = NULL, LONG32 priority = -2, LPCSTR ack = NULL, LPCSTR callback = NULL, LPCSTR value = NULL);
		LONG32 Notify(LPCWSTR classId = NULL, LPCWSTR title = NULL, LPCWSTR text = NULL, LONG32 timeout = -1, LPCWSTR iconPath = NULL, LPCWSTR iconBase64 = NULL, LONG32 priority = -2, LPCWSTR ack = NULL, LPCWSTR callback = NULL, LPCWSTR value = NULL);

		/// <summary>Register application with Snarl.</summary>
		/// <returns>The application token or negative on failure.</returns>
		/// <remarks>The application token is saved in SnarlInterface member variable, so just use return value to check for error.</remarks>
		LONG32 Register(LPCSTR  signature, LPCSTR  title, LPCSTR  icon = NULL, LPCSTR  password = NULL, HWND hWndReplyTo = NULL, LONG32 msgReply = 0);
		LONG32 Register(LPCWSTR signature, LPCWSTR title, LPCWSTR icon = NULL, LPCWSTR password = NULL, HWND hWndReplyTo = NULL, LONG32 msgReply = 0);

		/// <summary>Remove a notification class added with AddClass().</summary>
		LONG32 RemoveClass(LPCSTR classId);
		LONG32 RemoveClass(LPCWSTR classId);
		
		/// <summary>Update the text or other parameters of a visible Snarl notification.</summary>
		LONG32 Update(LONG32 msgToken, LPCSTR classId = NULL, LPCSTR title = NULL, LPCSTR text = NULL, LONG32 timeout = -1, LPCSTR iconPath = NULL, LPCSTR iconBase64 = NULL, LONG32 priority = -2, LPCSTR ack = NULL, LPCSTR callback = NULL, LPCSTR value = NULL);
		LONG32 Update(LONG32 msgToken, LPCWSTR classId = NULL, LPCWSTR title = NULL, LPCWSTR text = NULL, LONG32 timeout = -1, LPCWSTR iconPath = NULL, LPCWSTR iconBase64 = NULL, LONG32 priority = -2, LPCWSTR ack = NULL, LPCWSTR callback = NULL, LPCWSTR value = NULL);

		/// <summary>Unregister application with Snarl when application is closing.</summary>
		LONG32 Unregister(LPCSTR signature);
		LONG32 Unregister(LPCWSTR signature);

		/// <summary>Update information provided when calling RegisterApp.</summary>
		/*LONG32 UpdateApp(LPCSTR title = NULL, LPCSTR icon = NULL);
		LONG32 UpdateApp(LPCWSTR title = NULL, LPCWSTR icon = NULL);*/


	private:

		/// <summary>Convert a unicode string to UTF8</summary>
		/// <returns>Returns pointer to the new string - Remember to delete [] returned string !</returns>
		/// <remarks>Remember to call FreeString on returned string !!!</remarks>
		static LPSTR  WideToUTF8(LPCWSTR szWideStr);

		static LONG32 DoRequest(LPCSTR request, SnarlParameterList<char>& spl, UINT replyTimeout = 1000);
		static LONG32 DoRequest(LPCWSTR request, SnarlParameterList<wchar_t>& spl, UINT replyTimeout = 1000);

		void SetPassword(LPCSTR password);
		void SetPassword(LPCWSTR password);
		void ClearPassword();

		LONG32 appToken;
		LONG32 lastMsgToken;
		LPSTR szPasswordA;
		LPWSTR szPasswordW;
	}; // class SnarlInterface

	} // namespace V42
} // namespace Snarl

#endif // SNARL_INTERFACE_V42
