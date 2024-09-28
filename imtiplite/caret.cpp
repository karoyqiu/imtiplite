#include "stdafx.h"

#include <crtdbg.h>
#include <uiautomationclient.h>


namespace {

_COM_SMARTPTR_TYPEDEF(IUIAutomation, __uuidof(IUIAutomation));
_COM_SMARTPTR_TYPEDEF(IUIAutomationCacheRequest, __uuidof(IUIAutomationCacheRequest));
_COM_SMARTPTR_TYPEDEF(IUIAutomationElement, __uuidof(IUIAutomationElement));
_COM_SMARTPTR_TYPEDEF(IUIAutomationTextPattern2, __uuidof(IUIAutomationTextPattern2));
_COM_SMARTPTR_TYPEDEF(IUIAutomationTextRange, __uuidof(IUIAutomationTextRange));
_COM_SMARTPTR_TYPEDEF(IUIAutomationTextRangeArray, __uuidof(IUIAutomationTextRangeArray));
_COM_SMARTPTR_TYPEDEF(IAccessible, __uuidof(IAccessible));

IUIAutomationTextRangePtr GetTextRange(IUIAutomationElementPtr elem)
{
    IUIAutomationTextRangePtr range;
    IUIAutomationTextPattern2Ptr pattern;
    auto hr = elem->GetCachedPatternAs(UIA_TextPattern2Id, IID_PPV_ARGS(&pattern));

    if (pattern)
    {
        BOOL bActive = FALSE;
        hr = pattern->GetCaretRange(&bActive, &range);
    }
    else
    {
        hr = elem->GetCachedPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&pattern));

        if (pattern)
        {
            IUIAutomationTextRangeArrayPtr ranges;
            hr = pattern->GetSelection(&ranges);

            if (ranges)
            {
                int length = 0;
                hr = ranges->get_Length(&length);
                hr = ranges->GetElement(length - 1, &range);
            }
        }
    }

    return range;
}

BOOL GetCaretPosUIA(LPPOINT lpPoint)
{
    IUIAutomationPtr uia;
    auto hr = uia.CreateInstance(CLSID_CUIAutomation);

    IUIAutomationCacheRequestPtr cacheRequest;
    hr = uia->CreateCacheRequest(&cacheRequest);

    hr = cacheRequest->AddPattern(UIA_TextPatternId);
    hr = cacheRequest->AddPattern(UIA_TextPattern2Id);

    IUIAutomationElementPtr elem;
    hr = uia->GetFocusedElementBuildCache(cacheRequest, &elem);

    auto range = GetTextRange(elem);

    if (range)
    {
        hr = range->ExpandToEnclosingUnit(TextUnit_Character);

        SAFEARRAY *rects = nullptr;
        hr = range->GetBoundingRectangles(&rects);

        if (rects)
        {
            RECT *pRect = nullptr;
            int nRect = 0;
            hr = uia->SafeArrayToRectNativeArray(rects, &pRect, &nRect);
            SafeArrayDestroy(rects);

            if (nRect > 0)
            {
                _ASSERT(pRect);
                lpPoint->x = pRect->right;
                lpPoint->y = pRect->top;
                CoTaskMemFree(pRect);
                return TRUE;
            }
        }
    }

    return FALSE;
}

BOOL GetCaretPosMSAA(LPPOINT lpPoint)
{
    IAccessiblePtr acc;
    auto hWnd = GetForegroundWindow();
    auto dwThreadId = GetWindowThreadProcessId(hWnd, nullptr);

    GUITHREADINFO info = { 0 };
    info.cbSize = sizeof(info);
    GetGUIThreadInfo(dwThreadId, &info);

    auto hr = AccessibleObjectFromWindow(info.hwndFocus, static_cast<DWORD>(OBJID_CARET),
                                         IID_IAccessible, reinterpret_cast<void **>(&acc));

    if (acc)
    {
        long x = 0, y = 0, w = 0, h = 0;
        _variant_t child(CHILDID_SELF);
        child.vt = VT_I4;
        hr = acc->accLocation(&x, &y, &w, &h, child);

        if (hr == S_OK)
        {
            lpPoint->x = x + w;
            lpPoint->y = y;
            return TRUE;
        }
    }

    return FALSE;
}

}


BOOL GetCaretPosEx(LPPOINT lpPoint)
{
    lpPoint->x = 0;
    lpPoint->y = 0;

    return GetCaretPosUIA(lpPoint) || GetCaretPosMSAA(lpPoint);
}
