// SimpleTest.cpp
// This is quick test and example code for using the Snarl C++ Interface
// Includes basic Win32 example and STL example

#include "stdafx.h"

#include "..\..\..\SnarlInterface.h"
using namespace Snarl::V41;

#ifdef UNICODE
#define tout std::wcout
#else
#define tout std::cout
#endif

void Example1();
void Example2();


int _tmain(int argc, _TCHAR* argv[])
{
	if (!SnarlInterface::IsSnarlRunning()) {
		tout << _T("Snarl is not running") << std::endl;
		return 1;
	}

	Example1();
	Example2();

	return 0;
}

void Example1()
{
	SnarlInterface snarl;

	snarl.RegisterApp(_T("CppTest"), _T("C++ test app"), NULL);
	snarl.AddClass(_T("Class1"), _T("Class 1"));

	tout << _T("Ready for action. Will post some messages...") << std::endl;

	snarl.EZNotify(_T("Class1"), _T("C++ example 1"), _T("Some text"), 10);

	tout << _T("Hit a key to unregister") << std::endl;
	_getch();
	snarl.UnregisterApp();		
}

void Example2()
{
	SnarlInterface snarl;

	snarl.RegisterApp(_T("CppTest"), _T("C++ test app"), NULL);
	snarl.AddClass(_T("Class1"), _T("Class 1"));

	tout << _T("Ready for action. Will post some messages...") << std::endl;

	snarl.EZNotify(_T("Class1"), _T("C++ example 1"), _T("Some text"), 10);

	std::basic_stringstream<TCHAR> sstr1;
	sstr1 << _T("Size of TCHAR = ") << sizeof(TCHAR) << std::endl;
	sstr1 << _T("Snarl version = ") << snarl.GetVersion() << std::endl;
	sstr1 << _T("Snarl windows = ") << snarl.GetSnarlWindow() << std::endl;

	snarl.EZNotify(_T("Class1"), _T("Runtime info"), sstr1.str().c_str(), 10);
	sstr1 = std::basic_stringstream<TCHAR>();

	// -------------------------------------------------------------------

	// DON'T DO THIS
	// sstr1 << _T("Snarl icons path = ") << snarl.GetIconsPath() << std::endl;
	// We need to free the string!
		
	LPCTSTR tmp = snarl.GetIconsPath(); // Release with FreeString
	if (tmp != NULL) {
		sstr1 << tmp << _T("info.png");
		snarl.EZNotify(_T("Class1"), _T("Icon test"), _T("Some text and an icon"), 10, sstr1.str().c_str());
		snarl.FreeString(tmp);		
	}

	tout << _T("Hit a key to unregister") << std::endl;
	_getch();
	snarl.UnregisterApp();
}
