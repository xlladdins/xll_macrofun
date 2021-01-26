// xll_addin.cpp - information related to addins
//#define XLL_VERSION 4
#include "xll/xll/xll.h"

using namespace xll;

#define ARGS_HELP_URL "XLL.ARGS.html"

XLL_CONST(CSTRING, XLL_ARGS_PROCEDURE, _T("Procedure"),"Return \"Procedure\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_TYPE_TEXT, _T("TypeText"), "Return \"TypeText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_FUNCTION_TEXT, _T("FunctionText"), "Return \"FunctionText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_ARGUMENT_TEXT, _T("ArgumentText"), "Return \"ArgumentText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_MACRO_TYPE, _T("MacroType"), "Return \"MacroType\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_CATEGORY, _T("Category"), "Return \"Category\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_SHORTCUT_TEXT, _T("ShortcutText"), "Return \"ShortcutText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_HELP_TOPIC, _T("HelpTopic"), "Return \"HelpTopic\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_FUNCTION_HELP, _T("FunctionHelp"), "Return \"FunctionHelp\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_ARGUMENT_NAME, _T("ArgumentName"), "Return \"ArgumentName\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_ARGUMENT_HELP, _T("ArgumentHelp"), "Return \"ArgumentHelp\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_ARGUMENT_DEFAULT, _T("ArgumentDefault"), "Return \"ArgumentDefault\"", "XLL", ARGS_HELP_URL);

AddIn xai_args(
	Function(XLL_LPOPER, "xll_args", "XLL.ARGS")
	.Arguments({
		{XLL_LPOPER, "name", "is a function name or register id.", "XLL.ARGS"},
		{XLL_LPOPER, "keys", "is an array of keys from XLL_ARGS_*.", "{XLL_ARGS_PROCEDURE(), XLL_ARGS_FUNCTION_TEXT()}"},
	})
	.FunctionHelp("Return information about an add-in.")
	.Category("XLL")
	.Documentation(R"(
Return members of xll:Args for an add-in given its <c>name</c> or register id.
The <c>keys</c> are an array of strings. You can use <c>XLL_ARGS_*</c> to discover
known keys.
If called with no arguments return the default list of keys. If the second
argument is missing it will return the default keys of the first argument.
)")
);
LPOPER WINAPI xll_args(LPOPER pname, LPOPER pkeys)
{
#pragma XLLEXPORT
	// default keys
	static OPER keys({
		OPER("ModuleText"),
		OPER("Procedure"),
		OPER("TypeText"), 
		OPER("ArgumentText"),
		OPER("MacroType"),
		OPER("Category"),
		OPER("ShortcutText"),
		OPER("HelpTopic"),
		OPER("FunctionHelp"),
	});
	static OPER result;

	result = ErrNA;
	try {
		if (pname->xltype == xltypeMissing) {
			return &keys;
		}

		Args* pargs = nullptr;
		if (pname->is_str()) {
			pargs = AddIn::Arguments(*pname);
		}
		else if (pname->is_num()) {
			pargs = AddIn::Arguments(pname->as_num());
		}
		else {
			XLL_ERROR("XLL.ARGS: name must be a string or number");
		}
		ensure(pargs || !"XLL.ARGS: failed to find find function");

		if (pkeys->xltype == xltypeMissing) {
			pkeys = &keys;
		}
		result = OPER(rows(*pkeys), columns(*pkeys));
		for (unsigned i = 0; i < result.size(); ++i) {
			result[i] = (*pargs)[index(*pkeys, i)];
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("XLL.ADDINS: unknown exception");
	}

	return &result;
}

// ARGS.ARGUMENT("name"|"help"|"default")

AddIn xai_addins(
	Function(XLL_LPOPER, "xll_addins", "XLL.ADDINS")
	.FunctionHelp("Return array of all add-ins.")
	.Category("XLL")
	.Documentation(R"(
)")
);
LPOPER WINAPI xll_addins(void)
{
#pragma XLLEXPORT
	static OPER result;
	
	result = ErrNA;
	try {
		result = Nil;
		for (const auto& [key, args] : AddIn::Map) {
			OPER row = OPER({ key, args.RegisterId() });
			result.push_bottom(row);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("XLL.ADDINS: unknown exception");
	}

	return &result;
}