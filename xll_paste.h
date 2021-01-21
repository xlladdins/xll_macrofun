// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll/xll/xll.h"

namespace xll {

	// paste default argument at ref and return reference to what was pasted
	template<class X>
	inline XOPER<X> paste_default(XOPER<X> ref, const XArgs<X>* pa, unsigned i)
	{
		const XOPER<X>& x = pa->ArgumentDefault(i);

		if (x.is_str() and x.val.str[1] == '=') {
			// formula
			XOPER<X> xi = XExcel<X>(xlfEvaluate, x);
			ensure(xi);
			if (size(xi) == 1) {
				ensure(XExcel<X>(xlcFormula, x, ref));
			}
			else {
				auto rw = rows(xi);
				auto col = columns(xi);
				ref = XExcel<X>(xlfOffset, ref, XOPER<X>(0), XOPER<X>(0), XOPER<X>(rw), XOPER<X>(col));
				ensure(XExcel<X>(xlcFormulaArray, x, ref))
			}
		}
		else {
			auto rw = rows(x);
			auto col = columns(x);
			ref = XExcel<X>(xlfOffset, ref, XOPER<X>(0), XOPER<X>(0), XOPER<X>(rw), XOPER<X>(col));
			ensure(XExcel<X>(xlcFormula, x, ref));
		}

		return ref;
	}

}