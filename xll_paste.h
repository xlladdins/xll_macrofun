// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll/xll/xll.h"

namespace xll {

	using xll::Excel;

	// Paste argument at active cell and return a reference to what was pasted.
	// Default arguments are always strings. If they start with '='
	// then the result of evaluating the argument is pasted.
	// Arguments of the form '={a,...;b,...}' get pasted as arrays.
	// Arguments that are function calls are pasted as array formulas if necessary.
	inline OPER paste_default(const OPER& x)
	{
		OPER ref = Excel(xlfActiveCell);

		if (x.is_str() and x.val.str[1] == '=') {
			// formula
			OPER xi = Excel(xlfEvaluate, x);
			ensure(xi);
			if (size(xi) == 1) {
				ensure(Excel(xlcFormula, x, ref));
			}
			else {
				auto rw = rows(xi);
				auto col = columns(xi);
				ref = Excel(xlfOffset, ref, OPER(0), OPER(0), OPER(rw), OPER(col));
				ensure(Excel(xlcFormulaArray, x, ref));
			}
		}
		else {
			// paste as formula to evaluate the string
			ensure(Excel(xlcFormula, x, ref));
		}

		return ref;
	}

}