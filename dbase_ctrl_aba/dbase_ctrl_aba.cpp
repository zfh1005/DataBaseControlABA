// dbase_ctrl_aba.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#import "msado15.dll" no_namespace rename("EOF", "EndOfFile")
#define _dx_support_dbase_aba_dll_
#include "dbase_ctrl_aba.h"


#include <ole2.h>
#include <stdio.h>
#include <conio.h>

#include "objbase.h"
#include <string>
using namespace std;

struct dbase
{
	string strCmuName;
	int iStatus;
	dbase()
	{
		strCmuName="";
		iStatus=-1;
	}
};
/* global dbase control */
_ConnectionPtr pConnection = NULL;
_RecordsetPtr m_pRecordset;
char TableName[512]="";
string strLocalMcu=""; // store local computer name

#ifdef _MANAGED
#pragma managed(push, off)
#endif

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		CoInitialize(NULL);
		break;
	case DLL_PROCESS_DETACH:
		CoUninitialize();
		break;

	}
    return TRUE;
}

#ifdef _MANAGED
#pragma managed(pop)
#endif

bool dbConnect(char* pdbName, char* pdbTable, char* pComputerName)
{
	// inport COM
	if(pdbName == NULL || pdbTable == NULL || pComputerName==NULL)
	{
		return false;
	}

	//return true;
	strcpy_s(TableName,sizeof(TableName)-1,pdbTable);

	// Create dbase connection
	try
	{
		pConnection.CreateInstance(__uuidof(Connection));
		pConnection->ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=";
		pConnection->ConnectionString += pdbName;
		pConnection->ConnectionTimeout = 10;
		pConnection->Open("","","",adModeUnknown);	
		m_pRecordset.CreateInstance(__uuidof(Recordset));

	}
	catch(_com_error e)
	{
		printf(e.Description());
		printf("\nFAILED\n");
		return false;
	}
	printf("Database connection OK\n");	
	return true;
}

void dbDisconnect()
{
	//m_pRecordset->Close();
	m_pRecordset.Release();
	pConnection.Release();	
	if(pConnection != NULL)
	{
		pConnection->Close();
	}
	//pConnection.Release();
}

bool dbModify(char *pSn, int iStatus)
{
	if(!pConnection)
	{
		return false;
	}

	_bstr_t bstrSQL;
	bstrSQL="select * from ";
	bstrSQL+=TableName;
	bstrSQL+=" where DUT_SN='";
	bstrSQL+=pSn;
	bstrSQL+="'";
	HRESULT hr = m_pRecordset->Open(bstrSQL,pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	if(hr != S_OK)
	{
		if(m_pRecordset != NULL)
		{
			m_pRecordset->Close();
		}
		return false;
	}

	if(m_pRecordset->EndOfFile==0)
	{
		m_pRecordset->Close();
		char szStatus[4]="";
		sprintf_s(szStatus,sizeof(szStatus),"%d",iStatus);

		bstrSQL="update ";
		bstrSQL+=TableName;
		bstrSQL+=" set TEST_STATUS = ";
		bstrSQL+=szStatus;
		bstrSQL+=" where DUT_SN='";
		bstrSQL+=pSn;
		bstrSQL+="'";
		hr = m_pRecordset->Open(bstrSQL,pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
		if(hr != S_OK)
		{
			if(m_pRecordset != NULL)
			{
				m_pRecordset->Close();
			}
			return false;
		}
	}
	else
	{
		m_pRecordset->Close();
		char szStatus[4]="";
		sprintf_s(szStatus,sizeof(szStatus),"%d",iStatus);
		bstrSQL="insert into ";
		bstrSQL+=TableName;
		bstrSQL+="  (DUT_SN,ATE_NAME,TEST_STATUS) values ('";			
		bstrSQL+=pSn;
		bstrSQL+="','";
		bstrSQL+=strLocalMcu.c_str();
		bstrSQL+="','";
		bstrSQL+=szStatus;
		bstrSQL+="')";
		m_pRecordset->Open(bstrSQL,pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);		
		if(hr != S_OK)
		{
			if(m_pRecordset != NULL)
			{
				m_pRecordset->Close();
			}
			return false;
		}
	}

	return true;
}

int  dbQuery(char *pSn)
{
	_bstr_t bstrSQL;
	char szSQL[512]="";
	dbase data;
	string strCmuName="";
	string strStatus="";
	bstrSQL="select * from ";
	bstrSQL+=TableName;
	bstrSQL+="  where SN = '";			
	bstrSQL+=pSn;
	bstrSQL+="'";
	try
	{
		m_pRecordset->Open(bstrSQL,pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);

		if(m_pRecordset->EndOfFile==0)
		{
			data.strCmuName= (char*)(_bstr_t)m_pRecordset->GetCollect("ATN_NAME");
			data.iStatus = atoi((char*)(_bstr_t)m_pRecordset->GetCollect("TEST_STATUS"));
			m_pRecordset->MoveNext();
		}

		m_pRecordset->Close();
	}
	catch(_com_error e)
	{
		return false;
	}
	if(data.strCmuName == strLocalMcu  && data.iStatus == 1)
		return 0;       //A: FAIL  -> A:TESTING  -> Prompt "Change station to test"
	if(data.strCmuName == strLocalMcu && data.iStatus == 0)
		return 1;       //A: FAIL -> B:PASS
	if(data.strCmuName != strLocalMcu && data.iStatus == 0)
		return 3;		//A: FAIL -> B:PASS -> A: Not Testing, need return to A
	if(data.strCmuName != strLocalMcu && data.iStatus == 1)
		return 2;        //A:FAIL  -> B:TESTING	

	return -1;
}

bool dbDelete(char *pSn)
{
	_bstr_t bstrSQL;
	bstrSQL="delete * from ";
	bstrSQL+=TableName;
	bstrSQL+=" where DUT_SN='";
	bstrSQL+=pSn;
	bstrSQL+="'";

	HRESULT hr = m_pRecordset->Open(bstrSQL,pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	if(hr != S_OK)
	{
		if(m_pRecordset != NULL)
		{
			m_pRecordset->Close();
		}
		return false;
	}
	//m_pRecordset->Close();	

	return true;	
}