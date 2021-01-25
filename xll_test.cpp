// xll_test.cpp - test add-ins for paste function
#include "xll/xll/xll.h"

using namespace xll;

#ifdef _DEBUG

AddIn xai_test1(
	Function(XLL_DOUBLE, "xll_test1", "TEST1")
	.Args({
		{XLL_DOUBLE, "x", "is a num", "1.23"}
	})
);
double WINAPI xll_test1(double x)
{
#pragma XLLEXPORT
	return x;
}

AddIn xai_test2(
	Function(XLL_DOUBLE, "xll_test2", "TEST2")
	.Args({
		{XLL_LPOPER, "r", "is a range", "={1.23, 4.56}"} // 1x2
		})
);
double WINAPI xll_test2(LPOPER pr)
{
#pragma XLLEXPORT

	return Excel(xlfSum, *pr).val.num;
}

AddIn xai_test3(
	Function(XLL_DOUBLE, "xll_test3", "TEST3")
	.Args({
		{XLL_LPOPER, "r", "is a range", "={1.23; 4.56}"} // 2x1
	})
);
double WINAPI xll_test3(LPOPER pr)
{
#pragma XLLEXPORT

	return Excel(xlfSum, *pr).val.num;
}

AddIn xai_test4(
	Function(XLL_DOUBLE, "xll_test4", "TEST4")
	.Args({
		{XLL_LPOPER, "r", "is a range", "={1.23; 4.56}"}, // 2x1
		{XLL_LPOPER, "r2", "is another range", "={1,2;3,4}"}, // 2x2
		{XLL_CSTRING, "s", "is a string", "abc"}
	})
);
double WINAPI xll_test4(LPOPER pr, LPOPER pr2, traits<XLOPERX>::xcstr s)
{
#pragma XLLEXPORT

	return static_cast<double>(size(*pr) + size(*pr2) + _tcslen(s));
}

// TEST5 - return an array

#endif // _DEBUG