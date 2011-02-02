#ifndef SNARL_V42_TEST_HEADER
#define SNARL_V42_TEST_HEADER

#pragma once

#include "SnarlTestHelper.h"
#include "..\..\..\SnarlInterface_v42\SnarlInterface.h"

class CSnarlV42Test
{
public:
	void Test1();
	void Test2();
	void Test3();


	CSnarlV42Test(Snarl::V42::SnarlInterface* snarl, CSnarlTestHelper* pTestHelper);
	~CSnarlV42Test(void);

private:
	LPCTSTR GetIcon(int i);

	CSnarlTestHelper* pHelper;
	Snarl::V42::SnarlInterface* snarl;
	HWND hWnd;
};

#endif
