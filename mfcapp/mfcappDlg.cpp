// mfcappDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "mfcapp.h"
#include "mfcappDlg.h"
#include "CApplication.h"
#include "CBookmark0.h"
#include "CBookmarks.h"
#include "CDocument0.h"
#include "CDocuments.h"
#include "CSelection.h"
//#include "atltime.h "


#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CmfcappDlg* pDlg;

//xml 操作用
MSXML2::IXMLDOMDocument2Ptr   m_plDomDocument;  
MSXML2::IXMLDOMElementPtr   m_pDocRoot;

//word 操作用到
CApplication wordApp;
CDocuments docs;
CDocument0 docx;
CBookmarks bookmarks;
CBookmark0 bookmark;
CSelection selection;

//自定义函数
void ReadXML(CString XmlFilePath);
void FillBookmark(int id, int billItemType, CString billItemName, CString text);


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CmfcappDlg 对话框




CmfcappDlg::CmfcappDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CmfcappDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CmfcappDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CmfcappDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(BUTTON_PRINT, &CmfcappDlg::OnBnClickedPrint)
END_MESSAGE_MAP()


// CmfcappDlg 消息处理程序

BOOL CmfcappDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->ModifyStyleEx(CBS_SORT, 0, 0);
	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->AddString(_T("普通"));
	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->AddString(_T("次日到账"));
	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->SetCurSel(0);

	pDlg = this;

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CmfcappDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CmfcappDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
//
HCURSOR CmfcappDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CmfcappDlg::OnBnClickedPrint()
{
	// TODO: 在此添加控件通知处理程序代码

	TCHAR  pMoudlePath[MAX_PATH]; 
	CString wordFilePath;
	CString xmlFilePath;
	DWORD nPos;
	
	 nPos = GetCurrentDirectory( MAX_PATH, pMoudlePath);			//获取当前程序工作目录
	wordFilePath.Format (L"%s", pMoudlePath);
	xmlFilePath = wordFilePath;
	wordFilePath = wordFilePath + _T("\\WordAuto.docx");								//获取word文档的路径
	xmlFilePath = xmlFilePath  + _T("\\WordAuto.xml");								//获取xml文档的路径
	
	COleVariant varZero((short)0);
	COleVariant varTrue(short(1),VT_BOOL);
	COleVariant varFalse(short(0),VT_BOOL);
	COleVariant vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant varFilePath(wordFilePath);

	if (!wordApp.CreateDispatch(_T("Word.Application"), NULL))
    {
        AfxMessageBox(_T("本机没有安装word产品！"), MB_OK | MB_SETFOREGROUND);
        return;
    }
	wordApp.put_Visible(true);		//word.application测试时可见，隐藏请置为false


	docs = wordApp.get_Documents ();
	docx = docs.Open(varFilePath,				//FileName文件路径
											varFalse,					//ConfirmConversions确认转换
											varFalse, 					//ReadOnly只读
											varFalse,					//AddToRecentFiles添加到最近文件中
											vOptional,				//PasswordDocument文档口令
											vOptional,				//PasswordTemplate模版口令
											vOptional,				//Revert
											vOptional,				//WritePasswordDocument
											vOptional,				//WritePasswordTemplate
											vOptional,				//Format格式
											vOptional,				//Encoding编码
											vOptional,				//Visible可见
											vOptional,				//OpenAndRepair打开并修复
											vOptional,				//DocumentDirection
											vOptional,				//NoEncodingDialog
											vOptional				//XMLTransform
		);
	
	bookmarks = docx.get_Bookmarks ();

	ReadXML(xmlFilePath);

	//theApp.OnFilePrintSetup ();			//进行打印机设置

	//打印预览
	docx.PrintPreview();	
	//关闭打印预览
	docx.ClosePrintPreview();
	//打印,由于本机没安装打印机，提示错误，暂时注释
	//docx.PrintOut(varFalse,				// Background 
	//								vOptional,		// Append 
	//								vOptional,		// Range 
	//								vOptional,		// OutputFileName 
	//								vOptional,		// From 
	//								vOptional,		// To 
	//								vOptional,		// Item 
	//								vOptional,		// Copies 
	//								vOptional,		// Pages 
	//								vOptional,		// PageType 
	//								vOptional,		// PrintToFile 
	//								vOptional,		// Collate 
	//								vOptional,		// ActivePrinterMacGX 
	//								vOptional,		// ManualDuplexPrint
	//								vOptional,		// PrintZoomColumn
	//								vOptional,		// PrintZoomRow
	//								vOptional,		// PrintZoomPaperWidth
	//								vOptional		// PrintZoomPaperHeight
	//	);

	 // 退出word应用
	docx.Close(varFalse, vOptional, vOptional);
	wordApp.Quit(vOptional, vOptional, vOptional);
	wordApp.ReleaseDispatch();

	MessageBox(_T("完成"));
}

void ReadXML(CString XmlFilePath)
{
	int id;
	int billItemType;
	CString billItemName;
	CString text;

	 ::CoInitialize(NULL);  
	MSXML2::IXMLDOMDocumentPtr XMLDOC;   
	MSXML2::IXMLDOMElementPtr XMLROOT;  
	MSXML2::IXMLDOMElementPtr XMLELEMENT;  
	MSXML2::IXMLDOMNodeListPtr XMLNODES; //某个节点的所以字节点  
	MSXML2::IXMLDOMNamedNodeMapPtr XMLNODEATTS;//某个节点的所有属性;  
	MSXML2::IXMLDOMNodePtr XMLNODE;  
	HRESULT HR = XMLDOC.CreateInstance(_uuidof(MSXML2::DOMDocument30));  

	if(!SUCCEEDED(HR))  
	{
		MessageBox(NULL , _T("XML CreateInstance faild!!") , _T("Hint") , MB_OK);
		 return;  
	}  

	XMLDOC->load(COleVariant(XmlFilePath));  

	XMLROOT = XMLDOC->GetdocumentElement();//获得根节点;  
	XMLROOT->get_childNodes(&XMLNODES);//获得根节点的所有子节点;  

	long XMLNODESNUM,ATTSNUM;  
	XMLNODES->get_length(&XMLNODESNUM);//获得所有子节点的个数;  

	for(int i=0; i<XMLNODESNUM; i++)  
	{  
		 XMLNODES->get_item(i,&XMLNODE);//获得某个子节点;  
		 XMLNODE->get_attributes(&XMLNODEATTS);//获得某个节点的所有属性;  
		 XMLNODEATTS->get_length(&ATTSNUM);//获得所有属性的个数;  
		 for(int j=0; j<ATTSNUM; j++)  
		 {  
			  XMLNODEATTS->get_item(j,&XMLNODE);//获得某个属性;  
			  CString T1 = (LPCTSTR)(_bstr_t)XMLNODE->nodeName;  
			  CString T2 = (LPCTSTR)(_bstr_t)XMLNODE->text;  
			  if (T1 == "id"){
				  id = _ttoi(T2);
			  }else if(T1 == "billItemType")
			  {
				  billItemType = _ttoi(T2);
			  }else if(T1 == "billItemName")
			  {
				  billItemName = T2;
			  }else if(T1 == "text")
			  {
				  text = T2;
			  }else{
			  }
			  //MessageBox(NULL , (T1+" = "+T2) , _T("Hint") , MB_OK);
		 }  
		FillBookmark(id, billItemType, billItemName, text);
	}  

	XMLNODES.Release();  
	XMLNODE.Release();  
	XMLROOT.Release();  
	XMLDOC.Release();  
	::CoUninitialize(); 
}

void FillBookmark( int id, int billItemType, CString billItemName, CString text)
{
	CString str;
	int nSel;

	bookmark = bookmarks.Item (COleVariant(billItemName));
	bookmark.Select ();
	selection = wordApp.get_Selection();
	switch(billItemType)
	{
	case 0:			//textbox
			pDlg->GetDlgItem(id)->GetWindowText(str);
			selection.TypeText(str);
			break;
	case 1:			//radiobutton
			if (((CButton*)pDlg->GetDlgItem(id))->GetCheck()) {
				selection.TypeText(_T("√"));
			}else {
				selection.TypeText(_T(""));
			}
			break;
	case 2:			//combobox

			nSel = ( (CComboBox *)pDlg->GetDlgItem(id) ) ->GetCurSel();
			( (CComboBox *)pDlg->GetDlgItem(id) ) ->GetLBText(nSel, str);
			if (str == text) {
				selection.TypeText(_T("√"));
			}else {
				selection.TypeText(_T(""));
			}
			break;
	default:
			break;
	}
}