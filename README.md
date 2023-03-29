# xllutl
A thin wrapper of Excel XLL SDK, provide following utilities
1. CXLOPER12 class, it's the cpp extension of original SDK XLOPER12
2. xl12 function, convenient to call XLL C API
# pure SDK API vs xllutl
as a comparison, following is an XLL created from pure SDK API
```c++
#include "windows.h"
#include "xlcall.h"
/*
cl /nologo /std:c++latest /Zi /Zc:strictStrings- /ID:\excel2013sdk\include plainxll.cpp D:\excel2013sdk\src\xlcall.cpp /link /debug /dll /out:plainxll.xll /libpath:D:\excel2013sdk\lib\x64 xlcall32.lib
*/
extern "C" {
  
__declspec(dllexport)
LPXLOPER12 test() {
  auto ret = new XLOPER12;
  ret->xltype = xltypeInt | xlbitDLLFree;
  ret->val.w = 10;
  return ret;
}

void myxlFree(LPXLOPER12 px) {
  if (px->xltype == xltypeStr) {
    if (px->val.str) {
      delete px->val.str;
      px->val.str = nullptr;
    }
  } else if (px->xltype == xltypeRef ) {
    if(px->val.mref.lpmref) {
      delete px->val.mref.lpmref;
      px->val.mref.lpmref = nullptr;
    }
  } else if (px->xltype == xltypeMulti) {
    if (px->val.array.lparray) {
      auto count = px->val.array.rows*px->val.array.columns;
      for(unsigned i=0; i < count; i++) {
        myxlFree(&px->val.array.lparray[i]);
      }
      delete px->val.array.lparray;
      px->val.array.lparray = nullptr;
    }
  }
}
__declspec(dllexport)
void WINAPI xlAutoFree12(LPXLOPER12 pxFree) {
  myxlFree(pxFree);
  delete pxFree;
}

__declspec(dllexport)
int WINAPI xlAutoOpen() {
  XLOPER12 xExp;
  XLOPER12 xSig;
  XLOPER12 xFun;
  XLOPER12 xArg;
  XLOPER12 xMcr;
  XLOPER12 xCat;
  XLOPER12 xCut;
  XLOPER12 xTpc;
  XLOPER12 xHlp;
  XLOPER12 xBlk;
  
  xExp.xltype = 
  xSig.xltype = 
  xFun.xltype = 
  xArg.xltype = 
  xCat.xltype = 
  xCut.xltype = 
  xTpc.xltype = 
  xHlp.xltype = 
  xBlk.xltype = xltypeStr;
  xMcr.xltype = xltypeInt;
    
  xExp.val.str = L"\x0004test";
  xSig.val.str = L"\x0002Q$";
  xFun.val.str = L"\x0004test";
  xArg.val.str = L"\x0000";
  xMcr.val.w = 1;
  xCat.val.str = L"\x0004test";
  xCut.val.str = L"\x0000";
  xTpc.val.str = L"\x0000";
  xHlp.val.str = L"\x0000";
  xBlk.val.str = L"\x0001 ";  

  XLOPER12 xDll;
  XLOPER12 xRet;
  Excel12(xlGetName,&xDll,0);
  Excel12(xlfRegister,&xRet,11,&xDll,&xExp,&xSig,&xFun,&xArg,&xMcr,&xCat,&xCut,&xTpc,&xHlp,&xBlk);
  Excel12(xlFree,nullptr,1,&xDll);
  return 1;  
}
}
```
and following is the same XLL created with xllutl
```c++
#include "windows.h"
#include "xlcallex.h"
/*
cl /nologo /std:c++latest /Zi /Zc:strictStrings- /ID:\excel2013sdk\include plainxll.cpp xlcallex.cpp D:\excel2013sdk\src\xlcall.cpp /link /debug /dll /out:plainxll.xll /libpath:D:\excel2013sdk\lib\x64 xlcall32.lib user32.lib
*/
extern "C" {
  
__declspec(dllexport)
LPXLOPER12 test() {
  auto ret = new CXLOPER12(10);
  ret->dFree(true);
  return ret;
}

__declspec(dllexport)
void WINAPI xlAutoFree12(LPXLOPER12 pxFree) {
  delete static_cast<CXLOPER12*>(pxFree);
}

__declspec(dllexport)
int WINAPI xlAutoOpen() {
  auto xDll = xl12(xlGetName);
  xlfRegisterEx(&xDll,"test","Q$","test","",1,"test","","",""," ");
  return 1;  
}
}
```
have fun!
