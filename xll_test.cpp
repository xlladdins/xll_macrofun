// xll_test.cpp - test add-ins for paste function
#include "xll/xll/xll.h"
#include "xll_macrofun.h"

using namespace xll;

#ifdef _DEBUG

AddIn xai_alert(Macro("xll_alert", "XLL.ALERT"));
int WINAPI xll_alert()
{
#pragma XLLEXPORT
	OPER o = Excel(xlcAlert, OPER("Alert text"), OPER(1));

	return TRUE;
}

AddIn xai_object(Macro("xll_macro", "XLL.MACRO"));
int WINAPI xll_macro()
{
#pragma XLLEXPORT
	OPER A1(REF(0, 0));
	OPER B2(REF(1, 1));
	/*
	OPER z(0);
	OPER o = Excel(xlfCreateObject, OPER(7), A1, z, z, B2, z, z, OPER("*"));
	o = Excel(xlcAssignToObject, OPER("XLL.ALERT"));
	*/
	CreateObject button(CreateObject::Type::Button, A1, B2);
	button.text = "@";
	button.Object();
	button.Assign(OPER("XLL.ALERT"));

	//CreateObject dialog(CreateObject::Type::DialogFrame, OPER(REF(2, 2)), OPER(REF(6, 6)));
	//dialog.text = "Dialog Text";
	//dialog.Object();

	return TRUE;
}


inline bool test_map()
{
	{
		std::map<OPER, OPER> om;
		om[OPER("foo")] = "bar";
		ensure(om[OPER("foo")] == "bar");
		ensure(om[OPER("Foo")] == "bar");

		om[OPER("Foo")] = "baz";
		ensure(om[OPER("fOo")] == "baz");

		om[OPER("bam")] = OPER(2, 3);
		
		OPER& bam = om[OPER("bam")];
		ensure(rows(bam) == 2 and columns(bam) == 3);
		bam[3] = OPER(3, 2);
		bam[3][3] = "fiz";

		const OPER& bam2 = om[OPER("bam")];
		ensure(bam2[3] == bam2(1, 0));
		ensure(bam2[3][3] == bam2[3](1,1))
		ensure(bam2(1, 0)(1, 1) == "fiz");
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
		{XLL_LPOPER, "r", "is a range", "{1.23, 4.56}"} // 1x2
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
		{XLL_LPOPER, "r", "is a range", "{1.23; 4.56}"} // 2x1
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
		{XLL_LPOPER, "r", "is a range", "{1.23; 4.56}"}, // 2x1
		{XLL_LPOPER, "r2", "is another range", "{1,2;3,4}"}, // 2x2
		{XLL_CSTRING, "s", "is a string", "\"abc\""}
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
		{XLL_LPOPER, "r", "is a range", "{1.23; 4.56}"}, // 2x1
		{XLL_LPOPER, "r2", "is another range", "{1,2;3,4}"}, // 2x2
		{XLL_CSTRING, "s", "is a string", "\"abc\""}
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

// TEST5 - return an array
AddIn xai_test6(
	Function(XLL_LPOPER, "xll_test6", "TEST6")
	.Arguments({
		{XLL_DOUBLE, "d1", "is a double", "1.23"},
		{XLL_DOUBLE, "d2", "is a double", "RAND()"}, // 2x2
		})
);
LPOPER WINAPI xll_test6(double d1, double d2)
{
#pragma XLLEXPORT
	static OPER result;

	result = d1 + d2;

	return &result;
}
#endif // _DEBUG