// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll/xll/xll.h"

namespace xll {

	// Paste argument at active cell and return a reference to what was pasted.
	// Default arguments are always strings. If they start with '='
	// then the result of evaluating the argument is pasted.
	// Arguments of the form '={a,...;b,...}' get pasted as arrays.
	// Arguments that are function calls are pasted as array formulas if necessary.
	template<class X>
	inline XOPER<X> paste_default(const XOPER<X>& x)
	{
		XOPER<X> ref = XExcel<X>(xlfActiveCell);

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
				ensure(XExcel<X>(xlcFormulaArray, x, ref));
			}
		}
		else {
			// paste as formula to evaluate the string
			ensure(XExcel<X>(xlcFormula, x, ref));
		}

		return ref;
	}

}