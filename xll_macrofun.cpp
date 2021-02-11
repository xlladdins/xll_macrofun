// xll_macrofun.cpp - Sample xll project.
#include "xll_macrofun.h"

using namespace xll;

AddIn xai_get_document(
	Function(XLL_LPOPER, "xll_get_document", "XLL.GET.DOCUMENT")
	.Arguments({
		Arg(XLL_SHORT, "type_num", "is a number that specifies what type of information you want."),
		Arg(XLL_LPOPER, "name_text", "is the name of an open workbook."),
		})
	.Uncalced()
	.FunctionHelp("Returns information about a sheet in a workbook.")
	.Category("XLL")
	.HelpTopic("https://xlladdins.github.io/Excel4Macros/get.document.html")
);
LPOPER WINAPI xll_get_document(SHORT type_num, LPOPER pname_text)
{
#pragma XLLEXPORT
	static OPER o;
	
	try {
		XLOPER num = { .val = {.w = type_num}, .xltype = xltypeInt };
		o = Excel(xlfGetDocument, OPER(type_num), *pname_text);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_get_workspace(
	Function(XLL_LPOPER, "xll_get_workspace", "XLL.GET.WORKSPACE")
	.Arguments({
		Arg(XLL_WORD, "type_num", "is a number specifying the type of workspace information you want."),
		})
	.Uncalced()
	.FunctionHelp("Returns information about the workspace.")
	.Category("XLL")
	.HelpTopic("https://xlladdins.github.io/Excel4Macros/get.workspace.html")
);
LPOPER WINAPI xll_get_workspace(WORD type_num)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = Excel(xlfGetWorkspace, OPER(type_num));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}