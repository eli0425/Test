// mfcappDlg.cpp : ʵ���ļ�
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

//xml ������
MSXML2::IXMLDOMDocument2Ptr   m_plDomDocument;  
MSXML2::IXMLDOMElementPtr   m_pDocRoot;

//word �����õ�
CApplication wordApp;
CDocuments docs;
CDocument0 docx;
CBookmarks bookmarks;
CBookmark0 bookmark;
CSelection selection;

//�Զ��庯��
void ReadXML(CString XmlFilePath);
void FillBookmark(int id, int billItemType, CString billItemName, CString text);


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CmfcappDlg �Ի���




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


// CmfcappDlg ��Ϣ�������

BOOL CmfcappDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->ModifyStyleEx(CBS_SORT, 0, 0);
	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->AddString(_T("��ͨ"));
	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->AddString(_T("���յ���"));
	( (CComboBox *)GetDlgItem(COMBO_TYPE) ) ->SetCurSel(0);

	pDlg = this;

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CmfcappDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ��������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù����ʾ��
//
HCURSOR CmfcappDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CmfcappDlg::OnBnClickedPrint()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	TCHAR  pMoudlePath[MAX_PATH]; 
	CString wordFilePath;
	CString xmlFilePath;
	DWORD nPos;
	
	 nPos = GetCurrentDirectory( MAX_PATH, pMoudlePath);			//��ȡ��ǰ������Ŀ¼
	wordFilePath.Format (L"%s", pMoudlePath);
	xmlFilePath = wordFilePath;
	wordFilePath = wordFilePath + _T("\\WordAuto.docx");								//��ȡword�ĵ���·��
	xmlFilePath = xmlFilePath  + _T("\\WordAuto.xml");								//��ȡxml�ĵ���·��
	
	COleVariant varZero((short)0);
	COleVariant varTrue(short(1),VT_BOOL);
	COleVariant varFalse(short(0),VT_BOOL);
	COleVariant vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant varFilePath(wordFilePath);

	if (!wordApp.CreateDispatch(_T("Word.Application"), NULL))
    {
        AfxMessageBox(_T("����û�а�װword��Ʒ��"), MB_OK | MB_SETFOREGROUND);
        return;
    }
	wordApp.put_Visible(true);		//word.application����ʱ�ɼ�����������Ϊfalse


	docs = wordApp.get_Documents ();
	docx = docs.Open(varFilePath,				//FileName�ļ�·��
											varFalse,					//ConfirmConversionsȷ��ת��
											varFalse, 					//ReadOnlyֻ��
											varFalse,					//AddToRecentFiles��ӵ�����ļ���
											vOptional,				//PasswordDocument�ĵ�����
											vOptional,				//PasswordTemplateģ�����
											vOptional,				//Revert
											vOptional,				//WritePasswordDocument
											vOptional,				//WritePasswordTemplate
											vOptional,				//Format��ʽ
											vOptional,				//Encoding����
											vOptional,				//Visible�ɼ�
											vOptional,				//OpenAndRepair�򿪲��޸�
											vOptional,				//DocumentDirection
											vOptional,				//NoEncodingDialog
											vOptional				//XMLTransform
		);
	
	bookmarks = docx.get_Bookmarks ();

	ReadXML(xmlFilePath);

	//theApp.OnFilePrintSetup ();			//���д�ӡ������

	//��ӡԤ��
	docx.PrintPreview();	
	//�رմ�ӡԤ��
	docx.ClosePrintPreview();
	//��ӡ,���ڱ���û��װ��ӡ������ʾ������ʱע��
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

	 // �˳�wordӦ��
	docx.Close(varFalse, vOptional, vOptional);
	wordApp.Quit(vOptional, vOptional, vOptional);
	wordApp.ReleaseDispatch();

	MessageBox(_T("���"));
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
	MSXML2::IXMLDOMNodeListPtr XMLNODES; //ĳ���ڵ�������ֽڵ�  
	MSXML2::IXMLDOMNamedNodeMapPtr XMLNODEATTS;//ĳ���ڵ����������;  
	MSXML2::IXMLDOMNodePtr XMLNODE;  
	HRESULT HR = XMLDOC.CreateInstance(_uuidof(MSXML2::DOMDocument30));  

	if(!SUCCEEDED(HR))  
	{
		MessageBox(NULL , _T("XML CreateInstance faild!!") , _T("Hint") , MB_OK);
		 return;  
	}  

	XMLDOC->load(COleVariant(XmlFilePath));  

	XMLROOT = XMLDOC->GetdocumentElement();//��ø��ڵ�;  
	XMLROOT->get_childNodes(&XMLNODES);//��ø��ڵ�������ӽڵ�;  

	long XMLNODESNUM,ATTSNUM;  
	XMLNODES->get_length(&XMLNODESNUM);//��������ӽڵ�ĸ���;  

	for(int i=0; i<XMLNODESNUM; i++)  
	{  
		 XMLNODES->get_item(i,&XMLNODE);//���ĳ���ӽڵ�;  
		 XMLNODE->get_attributes(&XMLNODEATTS);//���ĳ���ڵ����������;  
		 XMLNODEATTS->get_length(&ATTSNUM);//����������Եĸ���;  
		 for(int j=0; j<ATTSNUM; j++)  
		 {  
			  XMLNODEATTS->get_item(j,&XMLNODE);//���ĳ������;  
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
				selection.TypeText(_T("��"));
			}else {
				selection.TypeText(_T(""));
			}
			break;
	case 2:			//combobox

			nSel = ( (CComboBox *)pDlg->GetDlgItem(id) ) ->GetCurSel();
			( (CComboBox *)pDlg->GetDlgItem(id) ) ->GetLBText(nSel, str);
			if (str == text) {
				selection.TypeText(_T("��"));
			}else {
				selection.TypeText(_T(""));
			}
			break;
	default:
			break;
	}
}