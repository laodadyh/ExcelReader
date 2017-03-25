#include <atlstr.h>
#include <vector>
#include <string>
#include <iostream>
using namespace std;

typedef std::vector < std::string > RowData;
typedef std::vector < RowData > ExcelData;

HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
{
	// Begin variable-argument list... 
	va_list marker;
	va_start(marker, cArgs);
	if (!pDisp) {
		cout << "NULL IDispatch passed to AutoWrap()";
		_exit(0);
	}
	// Variables used... 
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	char buf[200];
	char szName[200];
	// Convert down to ANSI 
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);
	// Get DISPID for name passed... 
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
		cout << buf << endl;
		_exit(0);
		return hr;
	}
	// Allocate memory for arguments... 
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	// Extract arguments... 
	for (int i = 0; i < cArgs; i++) {
		pArgs[i] = va_arg(marker, VARIANT);
	}
	// Build DISPPARAMS 
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;
	// Handle special-case for property-puts! 
	if (autoType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}
	// Make the call! 
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr)) {
		sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);
		cout << buf << endl;
		_exit(0);
		return hr;
	}
	// End variable-argument section... 
	va_end(marker);
	delete[] pArgs;
	return hr;
}

void read_excels(std::vector<std::string> excels, std::vector<ExcelData>& datas)
{
	CoInitialize(NULL);

	// 获得EXCEL的CLSID 
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

	if (FAILED(hr)) {
		cout << "CLSIDFromProgID() 函数调用失败!" << endl;
		return;
	}

	// 创建实例 
	IDispatch *pXlApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pXlApp);
	if (FAILED(hr)) {
		cout << "请检查是否已经安装EXCEL!";
		return;
	}

	// 显示，将Application.Visible属性置1 
	VARIANT x;
	x.vt = VT_I4;
	x.lVal = 0;
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x);
	// 获取Workbooks集合 
	IDispatch *pXlBooks;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0);
		pXlBooks = result.pdispVal;
	}

	//用来保存信息的数组 
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	SAFEARRAYBOUND sab[2];
	//	sab[0].lLbound = 1; sab[0].cElements = 40; 
	//	sab[1].lLbound = 1; sab[1].cElements = 16; 
	sab[0].lLbound = 1; sab[0].cElements = 100000;
	sab[1].lLbound = 1; sab[1].cElements = 500;
	arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
	int tableNum;

	for (auto& excel : excels)
	{
		ExcelData excel_data;
		CString strName = excel.c_str();  //Excel表完整路径 
		CString strTmp;                   //临时变量，保存单元格数据中的CString型 
		LONGLONG dblTmp;                  //临时变量，保存单元格数据中的int型

		//获取文件名
		CString str_temp;
		str_temp = strName.Left(strName.ReverseFind('.'));
		int index = str_temp.ReverseFind('/') > str_temp.ReverseFind('\\') ? str_temp.ReverseFind('/') : str_temp.ReverseFind('\\');
		str_temp = str_temp.Right(str_temp.GetLength() - index - 1);
		std::string node_name = str_temp.Right(str_temp.GetLength() - str_temp.ReverseFind('_') - 1);
		str_temp = str_temp.Left(str_temp.ReverseFind('_'));

		// 调用Workbooks.Open()方法，打开一个已经存在的Workbook 
		IDispatch *pXlBook;
		{
			VARIANT parm;
			parm.vt = VT_BSTR;
			// parm.bstrVal = ::SysAllocString(L"’’strName’’"); 
			parm.bstrVal = strName.AllocSysString();
			VARIANT result;
			VariantInit(&result);
			AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Open", 1, parm);
			pXlBook = result.pdispVal;
		}

		IDispatch *pXlSheet;
		{
			VARIANT result;
			VariantInit(&result);
			AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
			pXlSheet = result.pdispVal;
		}

		// 选择一个Range 
		IDispatch *pXlRange;
		{
			VARIANT parm;
			parm.vt = VT_BSTR;
			parm.bstrVal = ::SysAllocString(L"A1:EZ10000");

			VARIANT result;
			VariantInit(&result);
			AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
			VariantClear(&parm);

			pXlRange = result.pdispVal;
		}

		int m_size_column = 500;
		// 用这个Range读取数据 
		AutoWrap(DISPATCH_PROPERTYGET, &arr, pXlRange, L"Value", 0);
		bool is_unique[500] = { false };
		bool is_success = true;

		for (int i = 1; i <= 1000000; i++)
		{
			RowData row_data;
			for (int j = 1; j <= m_size_column; j++)
			{
				strTmp = "";
				VARIANT tmp;
				// 添加数据到数组中 
				long indices[] = { i, j };
				SafeArrayGetElement(arr.parray, indices, (void *)&tmp);
				if (tmp.vt == VT_BSTR)
				{
					strTmp = tmp.bstrVal;
				}
				else if (tmp.vt == VT_R8)
				{
					dblTmp = tmp.dblVal;
					strTmp.Format("%lld", dblTmp);
				}
				else if (tmp.vt == VT_NULL)
				{
					strTmp = "";
				}
				else
				{
					strTmp = "";
				}

				//最后一行
				if (j == 1 && strTmp.IsEmpty())
				{
					goto end;
				}
				//第一行最后一列
				if (i == 1 && strTmp.IsEmpty())
				{
					m_size_column = j;
					break;
				}
			}
		}
	end:
		AutoWrap(DISPATCH_METHOD, NULL, pXlBook, L"Close", 0);
		VariantClear(&arr);
		pXlRange->Release();
		pXlSheet->Release();
		pXlBook->Release();
	}

	// 退出，调用Application.Quit()方法 
	// 释放所有的接口以及变量 
	AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);
	pXlBooks->Release();
	pXlApp->Release();

	// 注销COM库 
	CoUninitialize();
}