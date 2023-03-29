#pragma once
#include <windows.h>
#include <strsafe.h>
#include <wchar.h>
#include <atomic>
#include <string>
#include <functional>
#include <type_traits>
#include "xlcall.h"

/*
Excel XLL utilities
2023/3/26
*/

enum class xltypeErrEx {
  NIL   = xlerrNull ,
  DIV0  = xlerrDiv0 ,
  VALUE = xlerrValue,
  REF   = xlerrRef  ,
  NAME  = xlerrName ,
  NUM   = xlerrNum  ,
  NA    = xlerrNA   ,
  GETTING_DATA = xlerrGettingData  ,
  MISSING = GETTING_DATA + 1,
};

template<unsigned N>
using refarray = std::array<XLREF12,N>;

struct CXLOPER12 : XLOPER12 {
  static CXLOPER12 nullop;
  static inline XLREF12 nullref;
  static inline std::atomic_int alloc = -1; // do not count nullop
  
  CXLOPER12() {
    xltype = xltypeNil;
    alloc++;
  }

  CXLOPER12(double d) {
    xltype = xltypeNum;
    val.num = d;
    alloc++;
  }

  CXLOPER12(int i) {
    xltype = xltypeInt;
    val.w = i;
    alloc++;
  }

  CXLOPER12(bool b) {
    xltype = xltypeBool;
    val.xbool = b;
    alloc++;
  }
  
  CXLOPER12(xltypeErrEx err) {
    xltype = (err != xltypeErrEx::MISSING ? xltypeErr : xltypeMissing);
    if(err != xltypeErrEx::MISSING)
      val.err = static_cast<int>(err);
    alloc++;
  }  
  // xltypeStr
  CXLOPER12(const char *str) {
    xltype = xltypeStr;
    size_t bytes;
    StringCbLengthA(str,STRSAFE_MAX_CCH * sizeof(TCHAR),&bytes);
    auto wlen= MultiByteToWideChar(CP_ACP,MB_ERR_INVALID_CHARS,str,bytes,nullptr,0);
    val.str = (XCHAR*)std::malloc(sizeof(wchar_t)*(wlen+1));
    val.str[0] = wlen;
    MultiByteToWideChar(CP_ACP,MB_ERR_INVALID_CHARS,str,bytes,val.str+1,wlen);
    alloc++;
  }
  CXLOPER12(const wchar_t *str) {
    xltype = xltypeStr;
    auto len = lstrlenW(str);
    val.str = (XCHAR*)std::malloc(sizeof(wchar_t)*(len+1));
    val.str[0] = len;
    wmemcpy_s(val.str+1, len, str, len);
    alloc++;
  } 
  // xltypeMulti
  CXLOPER12(RW r, COL c) {
    xltype = xltypeMulti;
    val.array.rows = r;
    val.array.columns = c;
    val.array.lparray = (LPXLOPER12)std::malloc(sizeof(CXLOPER12)*r*c);
    for(unsigned i=0 ; i<r*c; i++) {
      val.array.lparray[i].xltype = xltypeNil;
      alloc++;
    }
    alloc++;
  }
  
  CXLOPER12& at(RW r, COL c) {
    if (!isMulti() || r<1 || c<1 || r>val.array.rows || c>val.array.columns) {
      return (CXLOPER12&)nullop;
    } else {
      return (CXLOPER12&)val.array.lparray[(r-1)*val.array.columns+c-1];  
    }    
  }
  void each(std::function<bool(RW,COL,CXLOPER12&)> fn) {
    if (isMulti()) {
      for(RW r=0; r<val.array.rows; r++) {
        for(COL c=0; c<val.array.columns; c++) {
          if (!fn(r+1,c+1,(CXLOPER12&)val.array.lparray[r*val.array.columns+c]))
            goto exit;
        }
      }
    }
    exit:;
  }
  // xltypeSRef
  CXLOPER12(XLREF12 const&sref) {
    xltype = xltypeSRef;
    val.sref.count = 1;
    val.sref.ref = sref;
    alloc++;
  }
  // xltypeRef
  template<unsigned N>
  CXLOPER12(refarray<N> refs,IDSHEET sht) {
    xltype = xltypeRef;
    val.mref.idSheet = sht;
    auto bytes = sizeof(XLREF12)*N;
    auto lpmref = (XLMREF12*)std::malloc(bytes+sizeof(WORD));
    lpmref->count = N;
    memcpy(lpmref->reftbl,refs.data(),bytes);
    val.mref.lpmref = lpmref;
    alloc++;
  }
  
  XLREF12& at(unsigned idx) {
    if (!isRef() || idx<1 || idx>val.mref.lpmref->count) {
      return (XLREF12&)nullref;
    } else {
      return (XLREF12&)val.mref.lpmref[idx-1];
    }
  }
  void each(std::function<bool(unsigned,XLREF12&)> fn) {
    if(isRef()) {
      for(unsigned i=0; i<val.mref.lpmref->count; i++) {
        if(!fn(i+1,(XLREF12&)val.mref.lpmref->reftbl[i]))
          goto exit;
      }
    }
    exit:;
  }
  // not copyable
  CXLOPER12(CXLOPER12 &) = delete;
  CXLOPER12(const CXLOPER12 &) = delete;
  CXLOPER12(const volatile CXLOPER12 &) = delete;
  CXLOPER12& operator=(CXLOPER12&) = delete;
  CXLOPER12& operator=(const CXLOPER12&) = delete;
  CXLOPER12& operator=(const volatile CXLOPER12&) = delete;
  
  CXLOPER12(CXLOPER12 &&op) {
    myfree();
    move(op);
    alloc++;
  }
  CXLOPER12& operator=(CXLOPER12 &&op) {
    myfree();
    move(op);
    return*this;
  }
  static CXLOPER12& attach(LPXLOPER12 op) {
    return (CXLOPER12&)*op;
  }
  const char* type() {
    switch(xltype & 0xFFF) {
      case xltypeInt: {
        return "Int";
      }
      case xltypeNum: {
        return "Num";
      }
      case xltypeStr: {
        return "Str";
      }
      case xltypeBool: {
        return "Bool";
      }
      case xltypeRef: {
        return "Ref";
      }
      case xltypeErr: {
        return "Err";
      }
      case xltypeFlow: {
        return "Flow";
      }
      case xltypeMulti: {
        return "Multi";
      }
      case xltypeNil: {
        return "Nil";
      }
      case xltypeMissing: {
        return "Missing";
      }
      case xltypeSRef: {
        return "SRef";
      }
      case xltypeBigData: {
        return "BigData";
      }
      default: {
        return "";
      }
    }
  }
  bool isInt() {
    return (xltype & 0xFFF) == xltypeInt;
  }
  bool isNum() {
    return (xltype & 0xFFF) == xltypeNum;
  }
  bool isStr() {
    return (xltype & 0xFFF) == xltypeStr;
  }
  bool isBool() {
    return (xltype & 0xFFF) == xltypeBool;
  }
  bool isErr() {
    return (xltype & 0xFFF) == xltypeErr;
  }
  const char* err() {
    if (isErr()) {
      xltypeErrEx err = (xltypeErrEx)val.err;
      switch(err) {
        case xltypeErrEx::NIL: {
          return "#NULL!";
        }
        case xltypeErrEx::DIV0: {
          return "#DIV/0!";
        }
        case xltypeErrEx::VALUE: {
          return "#VALUE!";
        }
        case xltypeErrEx::REF: {
          return "#REF!";
        }
        case xltypeErrEx::NAME: {
          return "#NAME?";
        }
        case xltypeErrEx::NUM: {
          return "#NUM!";
        }
        case xltypeErrEx::NA: {
          return "#N/A";
        }
        case xltypeErrEx::GETTING_DATA: {
          return "#GETTING_DATA";
        }
        default: {
          return "";
        }
      }
    } else {
      return "";
    }
  }
  bool isFlow() {
    return (xltype & 0xFFF) == xltypeFlow;
  }
  bool isMulti() {
    return (xltype & 0xFFF) == xltypeMulti;
  }
  bool isNil() {
    return (xltype & 0xFFF) == xltypeNil;
  }
  bool isMissing() {
    return (xltype & 0xFFF) == xltypeMissing;
  }
  bool isRef() {
    return (xltype & 0xFFF) == xltypeRef;
  }
  bool isSRef() {
    return (xltype & 0xFFF) == xltypeSRef;
  }
  bool isBigData() {
    return (xltype & 0xFFF) == xltypeBigData;
  }
  operator bool() {
    return !isErr();
  }
  LPXLOPER12 operator &() {
    return this;
  }
  ~CXLOPER12() {
    myfree();
    alloc--;
    printf("~\n");
  }
  void dFree(bool flag) {
    if (flag) {
      xltype |= xlbitDLLFree;
    } else {
      xltype &= ~xlbitDLLFree;
    }
  }
  
  void xFree(bool flag) {
    if (flag) {
      xltype |= xlbitXLFree;
    } else {
      xltype &= ~xlbitXLFree;
    }
  }
private:
  /*
  why setup myfree
    if an UDF
    1. is not declared as thread safe ($) 
    2. and return a dynamic allocated XLOPER12 with xlbitDLLFree bit set
    calling xlFree in dtor will crash Excel.
    
    some times I would like to return dynamic allocated XLOPER12 with 
    xlbitDLLFree bit set in a non thread safe UDF, 
    for example,
    I can eliminate all uses of static XLOPER12
    myfree come into being.
  */
  void myfree() {
    if(isStr()) {
      if (val.str) {
        std::free(val.str);
        val.str = nullptr;
      }      
    } else if (isRef()) {
      if (val.mref.lpmref) {
        std::free(val.mref.lpmref);
        val.mref.lpmref = nullptr;  
      }      
    } else if (isMulti()) {
      if(val.array.lparray) {
        auto cells = val.array.rows*val.array.columns;
        for(unsigned i=0 ; i<cells; i++) {
          // mimic delete[]
          CXLOPER12 &op = (CXLOPER12&)val.array.lparray[cells-i-1];
          op.~CXLOPER12();
        }
        std::free(val.array.lparray);
        val.array.lparray = nullptr;
      }
    }
  }
  
  void move(CXLOPER12 &px) {
    memcpy(this,&px,sizeof(CXLOPER12));
    if(isStr()) {
      if(px.val.str) {
        px.val.str = nullptr;
      }
    } else if(isRef()) {
      if(px.val.mref.lpmref) {
        px.val.mref.lpmref = nullptr;
      }
    } else if(isMulti()) {
      if(px.val.array.lparray) {
        px.val.array.lparray = nullptr;  
      }      
    }
  }
};

static_assert(sizeof(CXLOPER12)==sizeof(XLOPER12));

template<typename ... ARGS>
requires std::conjunction_v<std::is_same<ARGS,LPXLOPER12>...>
[[nodiscard]]
CXLOPER12 xl12(unsigned xlfn, ARGS ... args) {
  auto count = sizeof ... (ARGS);
  CXLOPER12 xRet;
  Excel12(xlfn,&xRet,count,args...);
  return xRet;
}

template<typename ... ARGS>
[[nodiscard]]
CXLOPER12 xl12x(unsigned xlfn, ARGS ... args) {
  return xl12(xlfn,&CXLOPER12(args)...);
}

template<typename ...P>
requires std::conjunction_v<std::is_same<P,const char*>...>
void xlfRegisterEx(
  const char *fn,
  const char *fn_sig,
  const char *fn_name,
  const char *arg_names,
         int macro_type,
  const char *category,
  const char *shortcut,
  const char *topic,
  const char *fn_help,
  P...arg_helps)
{
  auto xDll = xl12(xlGetName);
  auto xRet = xl12(xlfRegister, 
    &xDll,
    &CXLOPER12(fn),
    &CXLOPER12(fn_sig),
    &CXLOPER12(fn_name),
    &CXLOPER12(arg_names),
    &CXLOPER12(macro_type),
    &CXLOPER12(category),
    &CXLOPER12(shortcut),
    &CXLOPER12(topic),
    &CXLOPER12(fn_help),
    &CXLOPER12(arg_helps)...);
}

struct xll {
  static bool called_from_wizard() {
    bool flag = false;
    auto callback = [](HWND hwnd,LPARAM lParam)->BOOL {
      bool *flag = (bool*)lParam;
      char classname[100] = {0};
      GetClassName(hwnd,classname,100);
      if (!strnicmp(classname,"bosa_sdm_xl",11)){
        *flag = true;
        return false; 
      }
      return true;
    };
    EnumWindows((WNDENUMPROC)callback,(LPARAM)&flag);
    return flag;
  }

  static std::string to_utf8(const char *str) {
    // convert to utf16
    int len,ret;
    len = MultiByteToWideChar(CP_ACP,MB_ERR_INVALID_CHARS,str,-1,nullptr,0);
    wchar_t *utf16str = new wchar_t[len];
    ret = MultiByteToWideChar(CP_ACP,MB_ERR_INVALID_CHARS,str,-1,utf16str,len);
    // convert utf16 to utf8
    len = WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,utf16str,-1,nullptr,0,nullptr,nullptr);
    char *utf8str = new char[len];
    ret = WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,utf16str,-1,utf8str,len,nullptr,nullptr);

    std::string tmp(utf8str);
    delete utf16str;
    delete utf8str;
    return tmp;
  }
};
// undocumented commands

#define xlcSetRec           (18 | xlCommand)
#define xlcRecMacro         (19 | xlCommand)
#define xlcAbsRelRec        (20 | xlCommand)
#define xlcSetRmPgBrk       (21 | xlCommand)

#define xlcPasteName        (58 | xlCommand)
#define xlcPasteFunc        (59 | xlCommand)
#define xlcRefTog           (60 | xlCommand)

#define xlcStartRecorder    (123 | xlCommand)

#define xlcHelpContents     (156 | xlCommand)
#define xlcHelpKeyboard     (157 | xlCommand)
#define xlcHelpLotus        (158 | xlCommand)

#define xlcHelpTutorial     (160 | xlCommand)

#define xlcInfoCell         (176 | xlCommand)
#define xlcInfoFmla         (177 | xlCommand)
#define xlcInfoValue        (178 | xlCommand)
#define xlcInfoFmt          (179 | xlCommand)
#define xlcInfoProt         (180 | xlCommand)
#define xlcInfoLbl          (181 | xlCommand)
#define xlcInfoPrec         (182 | xlCommand)
#define xlcInfoDep          (183 | xlCommand)
#define xlcInfoNote         (184 | xlCommand)

#define xlcGroupObj         (205 | xlCommand)

#define xlcUnused2          (221 | xlCommand)

#define xlcA230             (230 | xlCommand)
#define xlcA231             (231 | xlCommand)
#define xlcA232             (232 | xlCommand)
#define xlcA233             (233 | xlCommand)
#define xlcA234             (234 | xlCommand)
#define xlcA235             (235 | xlCommand)
#define xlcA236             (236 | xlCommand)
#define xlcA237             (237 | xlCommand)
#define xlcA238             (238 | xlCommand)
#define xlcA239             (239 | xlCommand)

#define xlcA241             (241 | xlCommand)
#define xlcA242             (242 | xlCommand)

#define xlcA244             (244 | xlCommand)
#define xlcA245             (245 | xlCommand)
#define xlcA246             (246 | xlCommand)
#define xlcA247             (247 | xlCommand)
#define xlcA248             (248 | xlCommand)

#define xlcSplitWindow      (255 | xlCommand)

#define xlcOLDPROJDisplayProjectSheet (257 | xlCommand)

#define xlcHelpBasic        (263 | xlCommand)
#define xlcHelpSwitch123    (264 | xlCommand)

#define xlcDoGal            (270 | xlCommand)
#define xlcHelpTroubleShooting (271 | xlCommand)

#define xlcA275             (275 | xlCommand)

#define xlcDisplayToolbar   (286 | xlCommand)
#define xlcHelpSearch       (287 | xlCommand)

#define xlcDeleteToolItem   (294 | xlCommand)

#define xlcShowToolbarPopup (299 | xlCommand)

#define xlcA303             (303 | xlCommand)
#define xlcA304             (304 | xlCommand)

#define xlcToolOption       (317 | xlCommand)

#define xlcUnused23         (326 | xlCommand)
#define xlcUnused24         (327 | xlCommand)

#define xlcCBTTableOfContents (329 | xlCommand)

#define xlcUnused27         (331 | xlCommand)
#define xlcUnused28         (332 | xlCommand)
#define xlcUnused29         (333 | xlCommand)
#define xlcUnused30         (334 | xlCommand)
#define xlcUnused22         (335 | xlCommand)

#define xlcFormatOption     (340 | xlCommand)
#define xlcFormatObj        (341 | xlCommand)

#define xlcUnused26         (345 | xlCommand)
#define xlcUnused31         (346 | xlCommand)
#define xlcUnused21         (347 | xlCommand)
#define xlcUnused33         (348 | xlCommand)
#define xlcUnused34         (349 | xlCommand)

#define xlcUnused6          (351 | xlCommand)

#define xlcCeFormatSeries   (357 | xlCommand)
#define xlcCeClearAll       (358 | xlCommand)
#define xlcCeClearVal       (359 | xlCommand)

#define xlcFormatChartPanel (371 | xlCommand)
#define xlcFormatCurSelection (372 | xlCommand)

#define xlcNewGal           (387 | xlCommand)

#define xlcInsNewWorkSheet  (401 | xlCommand)
#define xlcInsNewChart      (402 | xlCommand)
#define xlcInsNewOBModule   (403 | xlCommand)
#define xlcInsNewDialog     (404 | xlCommand)
#define xlcInsNewMacroSheet (405 | xlCommand)
#define xlcInsNewIntlMacro  (406 | xlCommand)
#define xlcToggleFmlabar    (407 | xlCommand)
#define xlcToggleStatusbar  (408 | xlCommand)
#define xlcInsertRows       (409 | xlCommand)
#define xlcInsertCols       (410 | xlCommand)
#define xlcChartWiz         (411 | xlCommand)

#define xlcCeErrbar         (418 | xlCommand)
#define xlcToggleSizeWWn    (419 | xlCommand)

#define xlcTracerIRefWhat   (426 | xlCommand)
#define xlcTracerIRefWhatD  (427 | xlCommand)
#define xlcTracerWhoRefsMe  (428 | xlCommand)
#define xlcTracerWhoRefsMeD (429 | xlCommand)

#define xlcCeTracerBar      (457 | xlCommand)

#define xlcCeHelpIndex      (479 | xlCommand)

#define xlcCeEditObjectConvert (483 | xlCommand)
#define xlcA484             (484 | xlCommand)

#define xlcHelpIntelliSearch (486 | xlCommand)
#define xlcHelpTFC          (487 | xlCommand)
#define xlcMergeCells       (488 | xlCommand)

#define xlcHelpOnKeyword    (490 | xlCommand)

#define xlcAlwaysCalc       (492 | xlCommand)

#define xlcUnused15         (496 | xlCommand)
#define xlcUnused16         (497 | xlCommand)
#define xlcUnused17         (498 | xlCommand)
#define xlcDrawInsert       (499 | xlCommand)
#define xlcDrawRepeat       (500 | xlCommand)
#define xlcInsPicEscher     (501 | xlCommand)
#define xlcInsPicServer     (502 | xlCommand)
#define xlcGBubble          (503 | xlCommand)
#define xlcCe3DShape        (504 | xlCommand)
#define xlcCDODataLabel     (505 | xlCommand)
#define xlcCDODataTable     (506 | xlCommand)
#define xlcUnused10         (507 | xlCommand)
#define xlcA508             (508 | xlCommand)

#define xlcCeVBAShowBookCode (512 | xlCommand)
#define xlcCeVBAShowSheetCode (513 | xlCommand)
#define xlcCeVBAShowIDE     (514 | xlCommand)
#define xlcCeVBAToggleMode  (515 | xlCommand)
#define xlcCeDisplayChartOptions (516 | xlCommand)
#define xlcA517             (517 | xlCommand)

#define xlcFreePgBrks       (524 | xlCommand)
#define xlcCeDataValidation (525 | xlCommand)
#define xlcCrtyp            (526 | xlCommand)
#define xlcCrtPlacement     (527 | xlCommand)
#define xlcGetExtData       (528 | xlCommand)
#define xlcModifyQryDef     (529 | xlCommand)
#define xlcModifyQryOpt     (530 | xlCommand)
#define xlcModifyQryParam   (531 | xlCommand)
#define xlcRefExtData       (532 | xlCommand)
#define xlcRefreshStatus    (533 | xlCommand)
#define xlcRefreshAll       (534 | xlCommand)
#define xlcHelpAboutXL      (535 | xlCommand)
#define xlcHelpMSN          (536 | xlCommand)
#define xlcInsertIndent     (537 | xlCommand)
#define xlcRowSeries        (538 | xlCommand)
#define xlcColSeries        (539 | xlCommand)
#define xlcCrtWiz2          (540 | xlCommand)
#define xlcSourceDataChart  (541 | xlCommand)
#define xlcClearSer         (542 | xlCommand)
#define xlcSeriesAddNew     (543 | xlCommand)
#define xlcCeClearChartData (544 | xlCommand)

#define xlcA550             (550 | xlCommand)
#define xlcMergeDocument    (551 | xlCommand)
#define xlcScaleTime        (552 | xlCommand)
#define xlcDataTablePatterns (553 | xlCommand)
#define xlcSetVPgBrk        (554 | xlCommand)
#define xlcSetHPgBrk        (555 | xlCommand)
#define xlcBubbleSizes      (556 | xlCommand)
#define xlcSeriesOptions    (557 | xlCommand)
#define xlcSxPivotCube      (558 | xlCommand)
#define xlcAccRejMark       (559 | xlCommand)

#define xlcAuditView        (560 | xlCommand)
#define xlcSxShowHeaderData (561 | xlCommand)
#define xlcSxShowHeader     (562 | xlCommand)
#define xlcSxShowData       (563 | xlCommand)
#define xlcSxShowTable      (564 | xlCommand)
#define xlcSxEnableSelection (565 | xlCommand)
#define xlcSxFormulaPane    (566 | xlCommand)
#define xlcSxProperties     (567 | xlCommand)
#define xlcSxConflict       (568 | xlCommand)
#define xlcSxConvertToBand  (569 | xlCommand)

#define xlcSxInsertField    (570 | xlCommand)
#define xlcSxInsertFormula  (571 | xlCommand)
#define xlcSxInsertItem     (572 | xlCommand)
#define xlcSxSelect         (573 | xlCommand)
#define xlcUseMe2           (574 | xlCommand)
#define xlcSxDrillDown      (575 | xlCommand)
#define xlcUseMe3           (576 | xlCommand)
#define xlcSxFieldAdvanced  (577 | xlCommand)
#define xlcA578             (578 | xlCommand)
#define xlcCancelRefresh    (579 | xlCommand)

#define xlcToggleChartWn    (580 | xlCommand)
#define xlcCrtwiz           (581 | xlCommand)
#define xlcToggleSecondaryPlot (582 | xlCommand)
#define xlcCondFmt          (583 | xlCommand)
#define xlcCeDValAudit      (584 | xlCommand)
#define xlcCeDValSelect     (585 | xlCommand)
#define xlcRunWebQuery      (586 | xlCommand)
#define xlcSxAutoFormat     (587 | xlCommand)
#define xlcCeRedo           (588 | xlCommand)
#define xlcCrtwizNewSh      (589 | xlCommand)

#define xlcCrtypDrp         (590 | xlCommand)
#define xlcCrtwizBtn        (591 | xlCommand)
#define xlcVbaHide          (592 | xlCommand)
#define xlcCeDValClearCircles (593 | xlCommand)
#define xlcCeDValMarkInvalid (594 | xlCommand)
#define xlcCeVBAShowPropertyBrowser (595 | xlCommand)
#define xlcCeInsertHyperlink (596 | xlCommand)
#define xlcNewInsTitle      (597 | xlCommand)
#define xlcCeSLVDragOff     (598 | xlCommand)
#define xlcCeHeaderFooter   (599 | xlCommand)

#define xlcRunDataQuery     (600 | xlCommand)
#define xlcCeHlinkOpen      (601 | xlCommand)
#define xlcUnused141        (602 | xlCommand)
#define xlcCeHlinkAddFav    (603 | xlCommand)
#define xlcCeHlinkEdit      (604 | xlCommand)
#define xlcCeHlinkEditCell  (605 | xlCommand)
#define xlcA606             (606 | xlCommand)
#define xlcFillEffects      (607 | xlCommand)
#define xlcGCylinder        (608 | xlCommand)
#define xlcGCone            (609 | xlCommand)

#define xlcGPyramid         (610 | xlCommand)
#define xlcCreateRenTask    (611 | xlCommand)
#define xlcMDIRestore       (612 | xlCommand)
#define xlcMDIMove          (613 | xlCommand)
#define xlcMDISize          (614 | xlCommand)
#define xlcMDIMinimize      (615 | xlCommand)
#define xlcMDIMaximize      (616 | xlCommand)
#define xlcCeHyperlink      (617 | xlCommand)
#define xlcCeClearHyperlink (618 | xlCommand)
#define xlcCeHelpContext    (619 | xlCommand)

#define xlcCePasteHlink     (622 | xlCommand)
#define xlcCeMSOnTheWeb1    (623 | xlCommand)
#define xlcCeMSOnTheWeb2    (624 | xlCommand)
#define xlcCeMSOnTheWeb3    (625 | xlCommand)
#define xlcCeMSOnTheWeb4    (626 | xlCommand)
#define xlcCeMSOnTheWeb5    (627 | xlCommand)
#define xlcCeMSOnTheWeb6    (628 | xlCommand)
#define xlcCeMSOnTheWeb7    (629 | xlCommand)

#define xlcCeMSOnTheWeb8    (630 | xlCommand)
#define xlcCeMSOnTheWeb9    (631 | xlCommand)
#define xlcCeMSOnTheWeb10   (632 | xlCommand)
#define xlcCeMSOnTheWeb11   (633 | xlCommand)
#define xlcCeMSOnTheWeb12   (634 | xlCommand)
#define xlcCeMSOnTheWeb13   (635 | xlCommand)
#define xlcCeMSOnTheWeb14   (636 | xlCommand)
#define xlcCeMSOnTheWeb15   (637 | xlCommand)
#define xlcCeMSOnTheWeb16   (638 | xlCommand)
#define xlcCeMSOnTheWeb17   (639 | xlCommand)

#define xlcCeToggleToolbars (640 | xlCommand)
#define xlcA641             (641 | xlCommand)
#define xlcA642             (642 | xlCommand)
#define xlcCrtGelProps      (643 | xlCommand)
#define xlcCrtGelPattern    (644 | xlCommand)
#define xlcCrtGelShaded     (645 | xlCommand)
#define xlcGenericRecorder  (646 | xlCommand)

#define xlcCtrlShftArrw     (648 | xlCommand)
#define xlcCtrlShftHome     (649 | xlCommand)
#define xlcRecHome          (650 | xlCommand)
#define xlcRecShiftHome     (651 | xlCommand)
#define xlcCtrlShftEnd      (652 | xlCommand)

#define xlcPhoneticInsert         (654 | xlCommand)
#define xlcPhoneticInsertSel      (655 | xlCommand)
#define xlcPhoneticProp           (656 | xlCommand)
#define xlcPhoneticShow           (657 | xlCommand)
#define xlcReconvertMSIME         (658 | xlCommand)
#define xlcHHConvert              (659 | xlCommand)
#define xlcGetMorphResult         (660 | xlCommand)
#define xlcCeSecurity             (661 | xlCommand)
#define xlcCeInsertScript         (662 | xlCommand)
#define xlcCeShowScriptAnchor     (663 | xlCommand)
#define xlcCeRemoveAllScripts     (664 | xlCommand)
#define xlcMicrosoftScriptEditor  (665 | xlCommand)
#define xlcImpTextFile            (666 | xlCommand)
