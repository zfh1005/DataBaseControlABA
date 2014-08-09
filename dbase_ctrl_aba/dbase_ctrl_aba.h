#ifndef _DBASE_CTRL_ABA_H_
#define _DBASE_CTRL_ABA_H_

#ifndef _dx_support_dbase_aba_dll_
#define DLLAPI_DBASE_ABA extern "C" __declspec(dllimport)
#else
#define DLLAPI_DBASE_ABA extern "C" __declspec(dllexport)
#endif

/* Init Access 2000/2003 database connection,*/
DLLAPI_DBASE_ABA bool dbConnect(char* pdbName/*database name, in*/, char* pdbTable/*Table name, in*/, char* pComputerName/*Local computer name, in*/);
/* Disconnect database connection, release resource*/
DLLAPI_DBASE_ABA void dbDisconnect();
/* Update the status of sn in databse */
DLLAPI_DBASE_ABA bool dbModify(char *pSn, int iStatus);
/* Query sn status:*/
/* 0: A: FAIL  -> A:TESTING  -> Prompt "Change station to test" */
/* 1: A: FAIL -> B:PASS */
/* 2: A:FAIL  -> B:TESTING	*/
/* 3: A: FAIL -> B:PASS -> A: Not Testing, need return to A */
DLLAPI_DBASE_ABA int  dbQuery(char *pSn);
/* Delete record by SN */
DLLAPI_DBASE_ABA bool dbDelete(char *pSn);

#endif