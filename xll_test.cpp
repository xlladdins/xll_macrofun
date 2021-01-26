// xll_test.cpp - test add-ins for paste function
#include "xll/xll/xll.h"

using namespace xll;

#ifdef _DEBUG

inline bool test_map()
{
	{
		std::map<OPER, OPER> om;
		om[OPER("foo")] = "bar";
		ensure(om[OPER("foo")] == "bar");
		ensure(om[OPER("Foo")] == "bar");

		om[OPER("Foo")] = "baz";
		ensure(om[OPER("fOo")] == "baz");
		int i;
		i = 1;
	}

	return true;
}
bool test_map_ = test_map();

AddIn xai_test1(
	Function(XLL_DOUBLE, "xll_test1", "TEST1")
	.Arguments({
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
	.Arguments({
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
	.Arguments({
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
	.Arguments({
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
AddIn xai_test5(
	Function(XLL_LPOPER, "xll_test5", "TEST5")
	.Arguments({
		{XLL_LPOPER, "r", "is a range", "={1.23; 4.56}"}, // 2x1
		{XLL_LPOPER, "r2", "is another range", "={1,2;3,4}"}, // 2x2
		{XLL_CSTRING, "s", "is a string", "\'abc"}
		})
);
LPOPER WINAPI xll_test5(LPOPER, LPOPER, traits<XLOPERX>::xcstr s)
{
#pragma XLLEXPORT
	static OPER result(3, 2);

	result(0, 0) = "third argument is:";
	result(2, 1) = OPER(s);

	return &result;
}

#endif // _DEBUG