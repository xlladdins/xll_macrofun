// spreadsheet.h - tools for automating spreadsheet creation
// https://xlladdins.github.io/Excel4Macros/
// 
// For xlcMacroFun having many arguments we define struct MacroFun.
// It has OPER members with exactly the same name as the documentation
// and a member OPER Fun() that calls Excel(xlcMacroFun, args...),
// and possibly come convenience functions making it easier to use.
// We throw in some enums from the documentation for the macro function
// and provide static members for common operations.
#pragma once
//#define XLL_VERSION 4
#include "xll/xll/xll.h"

namespace xll {

	using xll::Excel;

	//
	// Parameterized functions
	//

	// move relative to active cell
	template<class X = XLOPERX>
	inline XOPER<X> Move(int r, int c)
	{
		return XExcel<X>(xlcSelect,
			XExcel<X>(xlfOffset, XExcel<X>(xlfActiveCell), XOPER<X>(r), XOPER<X>(c))
		);
	}
	template<class X = XLOPERX>
	inline XOPER<X> Up(int r = 1)
	{
		return Move<X>(-r, 0);
	}
	template<class X = XLOPERX>
	inline XOPER<X> Down(int r = 1)
	{
		return Move<X>(r, 0);
	}
	template<class X = XLOPERX>
	inline XOPER<X> Right(int c = 1)
	{
		return Move<X>(0, c);
	}
	template<class X = XLOPERX>
	inline XOPER<X> Left(int c = 1)
	{
		return Move<X>(0, -c);
	}

	// RGB colors in the palette
	// https://xlladdins.github.io/Excel4Macros/edit.color.html
	class EditColor {
		inline static int count = 57; // largest index + 1
	public:
		OPER r, g, b;
		EditColor(int _r, int _g, int _b)
			: r(_r), g(_g), b(_b)
		{
			ensure(0 <= _r and _r <= 255);
			ensure(0 <= _g and _g <= 255);
			ensure(0 <= _b and _b <= 255);
		}
		// Add to color palette
		OPER Color() const
		{
			ensure(--count > 10);
			ensure(xll::Excel(xlcEditColor, OPER(count), OPER(r), OPER(g), OPER(b)));

			return OPER(count);
		}
		// Microsoft color palette
		static OPER MSred()
		{
			EditColor _(0xF6, 0x53, 0x14); return _.Color();
		}
		static OPER MSgreen()
		{
			EditColor _(0x7C, 0xBB, 0x00); return _.Color();
		}
		static OPER MSblue()
		{
			EditColor _(0x00, 0xA1, 0xF1); return _.Color();
		}
		static OPER MSorange()
		{
			EditColor _(0xFF, 0xBB, 0x00); return _.Color();
		}
	};

	// ALIGNMENT(horiz_align, wrap, vert_align, orientation, add_indent, shrink_to_fit, merge_cells)
	// https://xlladdins.github.io/Excel4Macros/alignment.html
	struct Alignment {
		OPER horiz_align = Missing;
		OPER wrap = Missing;
		OPER vert_align = Missing;
		OPER orientation = Missing;
		OPER add_indent = Missing;
		OPER shrink_to_fit = Missing;
		OPER merge_cells = Missing;

		enum class Horizontal {
			General = 1,
			Left,
			Center,
			Right,
			Fill,
			Justify,
			CenterCells
		};
		enum class Vertical {
			Top = 1,
			Center,
			Bottom,
			Justify
		};
		enum class Orientation {
			Horizontal = 0,
			Vertical,
			Upward,
			Downward
		};

		OPER Align()
		{
			return xll::Excel(xlcAlignment, horiz_align, wrap, vert_align, orientation);
		}
		static OPER Align(enum Horizontal e)
		{
			Alignment align; align.horiz_align = (int)e; return align.Align();
		}
		static OPER Align(enum Vertical e)
		{
			Alignment align; align.vert_align = (int)e; return align.Align();
		}
		static OPER Align(enum Orientation e)
		{
			Alignment align; align.orientation = (int)e; return align.Align();
		}
	};

	// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
	// https://xlladdins.github.io/Excel4Macros/format.font.html
	struct FormatFont {
		OPER name_text = Missing; // font name, e.g., "Calibri"
		OPER size_num = Missing;  // font size in points
		OPER bold = Missing;
		OPER italic = Missing;
		OPER underline = Missing;
		OPER strike = Missing;
		OPER color = Missing;
		OPER outline = Missing;
		OPER shadow = Missing;

		OPER Font()
		{
			return xll::Excel(xlcFormatFont, name_text, size_num,
				bold, italic, underline, strike,
				color, outline, shadow);
		}

		static OPER NameText(const char* name)
		{
			FormatFont ff; ff.name_text = name; return ff.Font();
		}
		static OPER SizeNum(int size)
		{
			FormatFont ff; ff.size_num = size; return ff.Font();
		}
		static OPER Color(const OPER& color)
		{
			FormatFont ff; ff.color = color; return ff.Font();
		}
		static OPER Bold(bool b = true)
		{
			FormatFont ff; ff.bold = b; return ff.Font();
		}
		static OPER Italic(bool b = true)
		{
			FormatFont ff; ff.italic = b; return ff.Font();
		}
		static OPER Underline(bool b = true)
		{
			FormatFont ff; ff.underline = b; return ff.Font();
		}
		static OPER Strike(bool b = true)
		{
			FormatFont ff; ff.strike = b; return ff.Font();
		}
	};

	// OPTIONS.CALCULATION(type_num, iter, max_num, max_change, update, precision, date_1904, calc_save, save_values)
	// https://xlladdins.github.io/Excel4Macros/options.calculation.html
	struct OptionsCalculation {
		enum class Type {
			Automatic = 1,
			AutomaticExceptTables = 2,
			Manual = 3
		};
		OPER type_num;
		OPER iter = OPER(false);
		OPER max_num = OPER(100);
		OPER max_change = OPER(0.001);
		OPER update = OPER(true);
		OPER precision = OPER(false);
		OPER date_1904 = OPER(false);
		OPER calc_save = OPER(true);
		OPER save_values = OPER(true);

		OptionsCalculation(enum Type type = Type::Automatic)
		{
			type_num = OPER((int)type);
		}

		OPER Calculation()
		{
			return xll::Excel(xlcOptionsCalculation, type_num, iter, max_num, max_change,
				update, precision, date_1904, calc_save, save_values);
		}
	};

	// OPTIONS.VIEW(formula, status, notes, show_info, object_num, page_breaks, formulas, gridlines, color_num, headers, outline, zeros, hor_scroll, vert_scroll, sheet_tabs)
	// https://xlladdins.github.io/Excel4Macros/options.view.html
	struct OptionsView {
		OPER formula = OPER(true);
		OPER status = OPER(true);
		OPER notes = OPER(false);
		OPER show_info = Missing;
		OPER object_num = Missing;
		OPER page_breaks = OPER(false);
		OPER formulas = OPER(false);
		OPER gridlines = OPER(true);
		OPER color_num = OPER(0);
		OPER headers = OPER(true);
		OPER outline = OPER(true);
		OPER zeros = OPER(true);
		OPER hor_scroll = OPER(true);
		OPER vert_scroll = OPER(true);
		OPER sheet_tabs = OPER(true);

		OPER View()
		{
			return xll::Excel(xlcOptionsView, formula, status, notes,
				show_info, object_num, page_breaks,
				formulas, gridlines, color_num, headers, outline,
				zeros, hor_scroll, vert_scroll, sheet_tabs);
		}

		OptionsView& Gridlines(bool b)
		{
			gridlines = b; return *this;
		}
	};


	// Cell information from xlcGetCell
	// https://xlladdins.github.io/Excel4Macros/get.cell.html
	struct Cell {
		OPER ref;
		Cell(const OPER& ref = xll::Excel(xlfActiveCell))
			: ref(ref)
		{ }
		Cell(REF ref)
			: ref(OPER(ref))
		{ }
		Cell(const char* ref)
			: ref(xll::Excel(xlfTextref, OPER(ref), OPER(false))) // R1C1
		{ }

		// generic get
		OPER Get(int i) const
		{
			return xll::Excel(xlfGetCell, OPER(i));
		}

		// Absolute reference of the upper-left cell in reference, as text in the current workspace reference style.
		OPER AbsRef() const
		{
			return xll::Excel(xlfGetCell, OPER(1), ref);
		}
		// Same as TYPE(reference).
		OPER Type() const
		{
			return xll::Excel(xlfType, ref);
		}
		// Contents of reference.
		OPER Contents() const
		{
			return xll::Excel(xlfGetCell, OPER(5), ref);
		}
		// Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.
		OPER Formula() const
		{
			return xll::Excel(xlfGetCell, OPER(6), ref);
		}
		// Number format of the cell
		OPER FormatNumber() const
		{
			return xll::Excel(xlfGetCell, OPER(7), ref);
		}
		Alignment Align() const
		{
			Alignment align;

			align.horiz_align = xll::Excel(xlfGetCell, OPER(7), ref);
			align.vert_align = xll::Excel(xlfGetCell, OPER(50), ref);
			align.orientation = xll::Excel(xlfGetCell, OPER(51), ref);

			return align;
		}
		// 14 	If the cell is locked, returns TRUE; otherwise, returns FALSE.
		// 15 	If the cell's formula is hidden, returns TRUE; otherwise, returns FALSE.
		// Row height of cell, in points.
		OPER RowHeight() const
		{
			return xll::Excel(xlfGetCell, OPER(17), ref);
		}
		// GET.CELL(44) - GET.CELL(42) to determine the width.
		OPER Width() const
		{
			return OPER(Get(44).val.num - Get(42).val.num);
		}
		// Name of font.
		OPER FontName() const
		{
			return xll::Excel(xlfGetCell, OPER(18), ref);
		}
		// Size of font, in points.
		OPER FontSize() const
		{
			return xll::Excel(xlfGetCell, OPER(19), ref);
		}
		// Style of the cell.
		OPER Style() const
		{
			return xll::Excel(xlfGetCell, OPER(40), ref);
		}
		// 46 	If the cell contains a text note, returns TRUE; otherwise, returns FALSE.
		// This returns the note text
		OPER Note() const
		{
			return xll::Excel(xlfGetNote, ref);
		}
		// Returns the name of the sheet in the form "[Book1]Sheet1".
		OPER Name() const
		{
			return xll::Excel(xlfGetCell, OPER(62), ref);
		}
	};

	// prefer styles to explicit formatting
	struct Style {
		OPER name;
		Style(const char* name)
			: name(name)
		{ }
		// https://xlladdins.github.io/Excel4Macros/apply.style.html
		OPER Apply() const
		{
			return xll::Excel(xlcApplyStyle, name);
		}
		static OPER Apply(const OPER& name)
		{
			return xll::Excel(xlcApplyStyle, name);
		}
		// https://xlladdins.github.io/Excel4Macros/delete.style.html
		OPER Delete() const 
		{
			return xll::Excel(xlcDeleteStyle, name);
		}
		static OPER Delete(const OPER& name)
		{
			return xll::Excel(xlcDeleteStyle, name);
		}
		// define styles
		Style& FormatNumber(const char* format)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(2), OPER(format)));

			return *this;
		}
		Style& FormatFont(const xll::FormatFont& format)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(3),
				format.name_text,
				format.size_num,
				format.bold,
				format.italic,
				format.underline,
				format.strike,
				format.color));

			return *this;
		}
		Style& Alignment(const xll::Alignment& align)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(4),
				align.horiz_align,
				align.wrap,
				align.vert_align,
				align.orientation));

			return *this;
		}
		//!!! 4-7
		Style& Pattern(const OPER& fore, const OPER& back = Missing)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(6), OPER(1), fore, back));

			return *this;
		}
	};

	// foreground and background cell colors
	// https://xlladdins.github.io/Excel4Macros/patterns.html
	inline OPER Patterns(OPER fore, OPER back)
	{
		return xll::Excel(xlcPatterns, OPER(1), fore, back);
	}

	// https://xlladdins.github.io/Excel4Macros/selection.html
	// https://xlladdins.github.io/Excel4Macros/select.html
	struct Select {
		OPER selection; // ??? must be an sref
		Select(const OPER& _selection = xll::Excel(xlfActiveCell))
			: selection(_selection)
		{
			if (selection.xltype & xltypeRef) {
				selection = OPER(selection.val.mref.lpmref->reftbl[0]);
			}

			ensure(Excel(xlcSelect, selection));
		}
		Select(const REF& selection)
			: Select(OPER(selection))
		{ }
		Select(const char* selection)
			: Select(xll::Excel(xlfTextref, OPER(selection), OPER(false))) // R1C1
		{ }

		OPER Offset(int r, int c, int h = 1, int w = 1) const
		{
			return xll::Excel(xlfOffset, selection, OPER(r), OPER(c), OPER(h), OPER(w));
		}

		// move r rows and c columns
		Select& Move(int r, int c)
		{
			selection = Offset(r, c);

			ensure(Excel(xlcSelect, selection));

			return *this;
		}
		Select& Down(int r = 1)
		{
			return Move(r, 0);
		}
		Select& Up(int r = 1)
		{
			return Move(-r, 0);
		}
		Select& Right(int c = 1)
		{
			return Move(0, c);
		}
		Select& Left(int c = 1)
		{
			return Move(0, -c);
		}

		// Enters numbers, text, references, and formulas in a worksheet
		OPER Formula(const OPER& formula)
		{
			return xll::Excel(xlcFormula, formula, selection);
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/xlset
		OPER Set(const OPER& set)
		{
			return xll::Excel(xlSet, selection, set);
		}
		OPER RowHeight(int points)
		{
			return xll::Excel(xlcRowHeight, OPER(points), selection);
		}
		OPER RowVisible(bool visible = true)
		{
			return xll::Excel(xlcRowHeight, OPER(), selection, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the row selection to a best-fit height
		OPER RowFit()
		{
			return xll::Excel(xlcRowHeight, OPER(), selection, OPER(), OPER(3));
		}
		OPER ColumnWidth(int points)
		{
			return xll::Excel(xlcColumnWidth, OPER(points), selection);
		}
		OPER ColumnVisible(bool visible = true)
		{
			return xll::Excel(xlcColumnWidth, OPER(), selection, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the column selection to a best-fit width
		OPER ColumnFit()
		{
			return xll::Excel(xlcColumnWidth, OPER(), selection, OPER(), OPER(3));
		}
		// Creates a comment 
		OPER Note(const char* note)
		{
			return xll::Excel(xlcNote, OPER(note), selection);
		}
		enum class Type {
			Notes = 1,
			Constants = 2,
			Formulas = 3,
			Blanks = 4,
			CurrentRegion = 5,
			Current = 6, // current array
			RowDifferences = 7,
			ColumnDifferences = 8,
			Precedents = 9,
			Dependents = 10,
			LastCell = 11,
			VisibleCellsOnly = 12, // (outlining)
			AllObjects = 13
		};
		enum class ValueType {
			Numbers = 1,
			Text = 2,
			Logical = 4,
			Error = 16,
			Default = Numbers + Text + Logical + Error
		};
		enum class Level {
			Direct = 1,
			All = 2
		};
		// uses active cell implicitly
		static OPER Special(Type type, ValueType value = ValueType::Default, Level level = Level::Direct)
		{
			OPER type_num = OPER((int)type);
			OPER value_type = Missing;
			OPER levels = Missing;

			if (type == Type::Constants or type == Type::Formulas) {
				value_type = OPER((int)value);
			}
			else if (type == Type::Precedents or type == Type::Dependents) {
				levels = OPER((int)level);
			}

			return xll::Excel(xlcSelectSpecial, type_num, value_type, levels);
		}
	};

	struct Name {
		OPER name;
		Name(const OPER& name)
			: name(name)
		{ }
		Name(const char* name)
			: name(name)
		{ }
		OPER Define(const OPER& ref, bool local)
		{
			return xll::Excel(xlcDefineName, name, ref, OPER(3), Missing, OPER(false), Missing, OPER(local));
		}
		OPER Delete()
		{
			return xll::Excel(xlcDeleteName, name);
		}
		OPER Get()
		{
			return xll::Excel(xlfGetName, name);
		}
		// xlcSetName is only for macro sheets
	};

	struct Document {
		static OPER Book()
		{
			return xll::Excel(xlfGetDocument, OPER(88)); // Book
		}

		static OPER Sheet()
		{
			return xll::Excel(xlfGetDocument, OPER(76)); // [Book]Sheet
		}
	};

	struct Workbook {
		static OPER Rename(const OPER& name)
		{
			return xll::Excel(xlcWorkbookName, Document::Book(), name);
		}
		static OPER Select(const OPER& name = Document::Sheet())
		{
			return xll::Excel(xlcWorkbookSelect, name);
		}
		// https://xlladdins.github.io/Excel4Macros/workbook.insert.html
		enum class Type {
			Worksheet = 1,
			Chart = 2,
			MacroSheet = 3,
			InternationallMacroSheet = 4,
			VisualBasicModule = 6,
			Dialog = 7
		};
		static OPER Insert(enum Type type = Type::Worksheet)
		{
			return xll::Excel(xlcWorkbookInsert, OPER((int)type));
		}
		static OPER Move(const OPER& name, int position = 1)
		{
			return xll::Excel(xlcWorkbookMove, name, Document::Sheet(), OPER(position));
		}
	};
}
