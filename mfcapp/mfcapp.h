// mfcapp.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CmfcappApp:
// �йش����ʵ�֣������ mfcapp.cpp
//

class CmfcappApp : public CWinApp
{
public:
	CmfcappApp();

// ��д
	public:
	virtual BOOL InitInstance();
	afx_msg void OnFilePrintSetup();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CmfcappApp theApp;