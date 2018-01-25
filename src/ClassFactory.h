//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa RichTextBox
//////////////////////////////////////////////////////////////////////


#pragma once

#include "OLEObject.h"
#include "OLEObjects.h"
#include "Table.h"
#include "TableRow.h"
#include "TableRows.h"
#include "TextFont.h"
#include "TextParagraph.h"
#include "TextParagraphTab.h"
#include "TextParagraphTabs.h"
#include "TextRange.h"
#include "TextSubRanges.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates a \c TextFont object</em>
	///
	/// Creates a \c TextFont object that represents a given font and color object and returns its
	/// \c IRichTextFont implementation as a smart pointer.
	///
	/// \param[in] pTextFont The \c ITextFont object for which to create the object.
	///
	/// \return A smart pointer to the created object's \c IRichTextFont implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774054.aspx">ITextFont</a>
	static inline CComPtr<IRichTextFont> InitTextFont(ITextFont* pTextFont)
	{
		CComPtr<IRichTextFont> pFont = NULL;
		InitTextFont(pTextFont, IID_IRichTextFont, reinterpret_cast<LPUNKNOWN*>(&pFont));
		return pFont;
	};

	/// \brief <em>Creates a \c TextFont object</em>
	///
	/// \overload
	///
	/// Creates a \c TextFont object that represents a given font and color object and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] pTextFont The \c ITextFont object for which to create the object.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppFont Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774054.aspx">ITextFont</a>
	static inline void InitTextFont(ITextFont* pTextFont, REFIID requiredInterface, LPUNKNOWN* ppFont)
	{
		ATLASSERT_POINTER(ppFont, LPUNKNOWN);
		ATLASSUME(ppFont);

		*ppFont = NULL;
		if(pTextFont) {
			// create a TextFont object and initialize it with the given parameters
			CComObject<TextFont>* pTextFontObj = NULL;
			CComObject<TextFont>::CreateInstance(&pTextFontObj);
			pTextFontObj->AddRef();
			pTextFontObj->Attach(pTextFont);
			pTextFontObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppFont));
			pTextFontObj->Release();
		}
	};

	/// \brief <em>Creates a \c TextParagraph object</em>
	///
	/// Creates a \c TextParagraph object that represents a given paragraph object and returns its
	/// \c IRichTextParagraph implementation as a smart pointer.
	///
	/// \param[in] pTextParagraph The \c ITextPara object for which to create the object.
	///
	/// \return A smart pointer to the created object's \c IRichTextParagraph implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>
	static inline CComPtr<IRichTextParagraph> InitTextParagraph(ITextPara* pTextParagraph)
	{
		CComPtr<IRichTextParagraph> pParagraph = NULL;
		InitTextParagraph(pTextParagraph, IID_IRichTextParagraph, reinterpret_cast<LPUNKNOWN*>(&pParagraph));
		return pParagraph;
	};

	/// \brief <em>Creates a \c TextParagraph object</em>
	///
	/// \overload
	///
	/// Creates a \c TextParagraph object that represents a given paragraph object and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] pTextParagraph The \c ITextPara object for which to create the object.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppParagraph Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>
	static inline void InitTextParagraph(ITextPara* pTextParagraph, REFIID requiredInterface, LPUNKNOWN* ppParagraph)
	{
		ATLASSERT_POINTER(ppParagraph, LPUNKNOWN);
		ATLASSUME(ppParagraph);

		*ppParagraph = NULL;
		if(pTextParagraph) {
			// create a TextParagraph object and initialize it with the given parameters
			CComObject<TextParagraph>* pTextParagraphObj = NULL;
			CComObject<TextParagraph>::CreateInstance(&pTextParagraphObj);
			pTextParagraphObj->AddRef();
			pTextParagraphObj->Attach(pTextParagraph);
			pTextParagraphObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppParagraph));
			pTextParagraphObj->Release();
		}
	};

	/// \brief <em>Creates a \c TextParagraphTab object</em>
	///
	/// Creates a \c TextParagraphTab object that represents one specific tab in a paragraph and returns its
	/// \c IRichTextParagraphTab implementation as a smart pointer.
	///
	/// \param[in] tabPosition The position of the tab for which to create the object.
	/// \param[in] pTextParagraph The \c ITextPara object the \c TextParagraphTab object will work on.
	/// \param[in] validateTabPosition If \c TRUE, the method will check whether the \c tabPosition parameter
	///            identifies an existing tab; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IRichTextParagraphTab implementation.
	static inline CComPtr<IRichTextParagraphTab> InitTextParagraphTab(float tabPosition, ITextPara* pTextParagraph, BOOL validateTabPosition = TRUE)
	{
		CComPtr<IRichTextParagraphTab> pTab = NULL;
		InitTextParagraphTab(tabPosition, pTextParagraph, IID_IRichTextParagraphTab, reinterpret_cast<LPUNKNOWN*>(&pTab), validateTabPosition);
		return pTab;
	};

	/// \brief <em>Creates a \c TextParagraphTab object</em>
	///
	/// \overload
	///
	/// Creates a \c TextParagraphTab object that represents one specific tab in a paragraph and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] tabPosition The position of the tab for which to create the object.
	/// \param[in] pTextParagraph The \c ITextPara object the \c TextParagraphTab object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTab Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateTabPosition If \c TRUE, the method will check whether the \c tabPosition parameter
	///            identifies an existing tab; otherwise not.
	static inline void InitTextParagraphTab(float tabPosition, ITextPara* pTextParagraph, REFIID requiredInterface, LPUNKNOWN* ppTab, BOOL validateTabPosition = TRUE)
	{
		ATLASSERT_POINTER(ppTab, LPUNKNOWN);
		ATLASSUME(ppTab);

		*ppTab = NULL;
		if(validateTabPosition && !IsValidTextParagraphTabPosition(tabPosition, pTextParagraph)) {
			// there's nothing useful we could return
			return;
		}

		// create a TextParagraphTab object and initialize it with the given parameters
		CComObject<TextParagraphTab>* pTextParagraphTabObj = NULL;
		CComObject<TextParagraphTab>::CreateInstance(&pTextParagraphTabObj);
		pTextParagraphTabObj->AddRef();
		pTextParagraphTabObj->Attach(pTextParagraph, tabPosition);
		pTextParagraphTabObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTab));
		pTextParagraphTabObj->Release();
	};

	/// \brief <em>Creates a \c TextParagraphTabs object</em>
	///
	/// Creates a \c TextParagraphTabs object that represents a collection of a paragraph's tabs and returns
	/// its \c IRichTextParagraphTabs implementation as a smart pointer.
	///
	/// \param[in] pTextParagraph The \c ITextPara object the \c TextParagraphTabs object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTextParagraphTabs implementation.
	static inline CComPtr<IRichTextParagraphTabs> InitTextParagraphTabs(ITextPara* pTextParagraph)
	{
		CComPtr<IRichTextParagraphTabs> pTabs = NULL;
		InitTextParagraphTabs(pTextParagraph, IID_IRichTextParagraphTabs, reinterpret_cast<LPUNKNOWN*>(&pTabs));
		return pTabs;
	};

	/// \brief <em>Creates a \c TextParagraphTabs object</em>
	///
	/// \overload
	///
	/// Creates a \c TextParagraphTabs object that represents a collection of a paragraph's tabs and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pTextParagraph The \c ITextPara object the \c TextParagraphTabs object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTabs Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitTextParagraphTabs(ITextPara* pTextParagraph, REFIID requiredInterface, LPUNKNOWN* ppTabs)
	{
		ATLASSERT_POINTER(ppTabs, LPUNKNOWN);
		ATLASSUME(ppTabs);

		*ppTabs = NULL;
		// create a TextParagraphTabs object and initialize it with the given parameters
		CComObject<TextParagraphTabs>* pTextParagraphTabsObj = NULL;
		CComObject<TextParagraphTabs>::CreateInstance(&pTextParagraphTabsObj);
		pTextParagraphTabsObj->AddRef();
		pTextParagraphTabsObj->Attach(pTextParagraph);
		pTextParagraphTabsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTabs));
		pTextParagraphTabsObj->Release();
	};

	/// \brief <em>Creates a \c TextRange object</em>
	///
	/// Creates a \c TextRange object that represents a given text range and returns its \c IRichTextRange
	/// implementation as a smart pointer.
	///
	/// \param[in] pTextRange The \c ITextRange object for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TextRange object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTextRange implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline CComPtr<IRichTextRange> InitTextRange(ITextRange* pTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTextRange> pRange = NULL;
		InitTextRange(pTextRange, pRichTextBox, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(&pRange));
		return pRange;
	};

	/// \brief <em>Creates a \c TextRange object</em>
	///
	/// \overload
	///
	/// Creates a \c TextRange object that represents a given text range and returns its implementation of a
	/// given interface as a raw pointer.
	///
	/// \param[in] pTextRange The \c ITextRange object for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TextRange object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppRange Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline void InitTextRange(ITextRange* pTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppRange)
	{
		ATLASSERT_POINTER(ppRange, LPUNKNOWN);
		ATLASSUME(ppRange);

		*ppRange = NULL;
		if(pTextRange) {
			// create a TextRange object and initialize it with the given parameters
			CComObject<TextRange>* pTextRangeObj = NULL;
			CComObject<TextRange>::CreateInstance(&pTextRangeObj);
			pTextRangeObj->AddRef();
			pTextRangeObj->SetOwner(pRichTextBox);
			pTextRangeObj->Attach(pTextRange);
			pTextRangeObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppRange));
			pTextRangeObj->Release();
		}
	};

	/// \brief <em>Creates a \c TextSubRanges object</em>
	///
	/// Creates a \c TextSubRanges object that represents a collection of text ranges that are sub-ranges of
	/// another text range. It returns the collection's \c IRichTextSubRanges implementation as a smart
	/// pointer.
	///
	/// \param[in] pTextRange The \c ITextRange implementation of the parent text range.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TextSubRanges object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTextSubRanges implementation.
	static inline CComPtr<IRichTextSubRanges> InitTextSubRanges(ITextRange* pTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTextSubRanges> pSubRanges = NULL;
		InitTextSubRanges(pTextRange, pRichTextBox, IID_IRichTextSubRanges, reinterpret_cast<LPUNKNOWN*>(&pSubRanges));
		return pSubRanges;
	};

	/// \brief <em>Creates a \c TextSubRanges object</em>
	///
	/// \overload
	///
	/// Creates a \c TextSubRanges object that represents a collection of text ranges that are sub-ranges of
	/// another text range. It returns the collection's \c IRichTextSubRanges implementation as a smart
	/// pointer.
	///
	/// \param[in] pTextRange The \c ITextRange implementation of the parent text range.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TextSubRanges object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppSubRanges Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitTextSubRanges(ITextRange* pTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppSubRanges)
	{
		ATLASSERT_POINTER(ppSubRanges, LPUNKNOWN);
		ATLASSUME(ppSubRanges);

		*ppSubRanges = NULL;
		if(pTextRange) {
			// create a TextSubRanges object and initialize it with the given parameters
			CComObject<TextSubRanges>* pTextSubRangesObj = NULL;
			CComObject<TextSubRanges>::CreateInstance(&pTextSubRangesObj);
			pTextSubRangesObj->AddRef();
			pTextSubRangesObj->SetOwner(pRichTextBox);
			pTextSubRangesObj->Attach(pTextRange);
			pTextSubRangesObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppSubRanges));
			pTextSubRangesObj->Release();
		}
	};

	/// \brief <em>Creates a \c Table object</em>
	///
	/// Creates a \c Table object that represents a given text range and returns its \c IRichTable
	/// implementation as a smart pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c Table object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTable implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline CComPtr<IRichTable> InitTable(ITextRange* pTableTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTable> pTable = NULL;
		InitTable(pTableTextRange, pRichTextBox, IID_IRichTable, reinterpret_cast<LPUNKNOWN*>(&pTable));
		return pTable;
	};

	/// \brief <em>Creates a \c Table object</em>
	///
	/// \overload
	///
	/// Creates a \c Table object that represents a given text range and returns its implementation of a
	/// given interface as a raw pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c Table object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTable Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline void InitTable(ITextRange* pTableTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppTable)
	{
		ATLASSERT_POINTER(ppTable, LPUNKNOWN);
		ATLASSUME(ppTable);

		*ppTable = NULL;
		if(pTableTextRange) {
			// create a Table object and initialize it with the given parameters
			CComObject<Table>* pTableObj = NULL;
			CComObject<Table>::CreateInstance(&pTableObj);
			pTableObj->AddRef();
			pTableObj->SetOwner(pRichTextBox);
			pTableObj->Attach(pTableTextRange);
			pTableObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTable));
			pTableObj->Release();
		}
	};

	/// \brief <em>Creates a \c TableRow object</em>
	///
	/// Creates a \c TableRow object that represents a given text range and returns its \c IRichTableRow
	/// implementation as a smart pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableRow object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTableRow implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline CComPtr<IRichTableRow> InitTableRow(ITextRange* pTableTextRange, ITextRange* pRowTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTableRow> pTableRow = NULL;
		InitTableRow(pTableTextRange, pRowTextRange, pRichTextBox, IID_IRichTableRow, reinterpret_cast<LPUNKNOWN*>(&pTableRow));
		return pTableRow;
	};

	/// \brief <em>Creates a \c TableRow object</em>
	///
	/// \overload
	///
	/// Creates a \c TableRow object that represents a given text range and returns its implementation of a
	/// given interface as a raw pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableRow object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTableRow Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline void InitTableRow(ITextRange* pTableTextRange, ITextRange* pRowTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppTableRow)
	{
		ATLASSERT_POINTER(ppTableRow, LPUNKNOWN);
		ATLASSUME(ppTableRow);

		*ppTableRow = NULL;
		if(pTableTextRange && pRowTextRange) {
			// create a TableRow object and initialize it with the given parameters
			CComObject<TableRow>* pTableRowObj = NULL;
			CComObject<TableRow>::CreateInstance(&pTableRowObj);
			pTableRowObj->AddRef();
			pTableRowObj->SetOwner(pRichTextBox);
			pTableRowObj->Attach(pTableTextRange, pRowTextRange);
			pTableRowObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTableRow));
			pTableRowObj->Release();
		}
	};

	/// \brief <em>Creates a \c TableRows object</em>
	///
	/// Creates a \c TableRows object that represents a given text range and returns its \c IRichTableRows
	/// implementation as a smart pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableRows object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTableRows implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline CComPtr<IRichTableRows> InitTableRows(ITextRange* pTableTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTableRows> pTableRows = NULL;
		InitTableRows(pTableTextRange, pRichTextBox, IID_IRichTableRows, reinterpret_cast<LPUNKNOWN*>(&pTableRows));
		return pTableRows;
	};

	/// \brief <em>Creates a \c TableRows object</em>
	///
	/// \overload
	///
	/// Creates a \c TableRows object that represents a given text range and returns its implementation of a
	/// given interface as a raw pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableRows object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTableRows Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline void InitTableRows(ITextRange* pTableTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppTableRows)
	{
		ATLASSERT_POINTER(ppTableRows, LPUNKNOWN);
		ATLASSUME(ppTableRows);

		*ppTableRows = NULL;
		if(pTableTextRange) {
			// create a TableRows object and initialize it with the given parameters
			CComObject<TableRows>* pTableRowsObj = NULL;
			CComObject<TableRows>::CreateInstance(&pTableRowsObj);
			pTableRowsObj->AddRef();
			pTableRowsObj->SetOwner(pRichTextBox);
			pTableRowsObj->Attach(pTableTextRange);
			pTableRowsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTableRows));
			pTableRowsObj->Release();
		}
	};

	/// \brief <em>Creates a \c TableCell object</em>
	///
	/// Creates a \c TableCell object that represents a given text range and returns its \c IRichTableCell
	/// implementation as a smart pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row for which to create the object.
	/// \param[in] pCellTextRange The \c ITextRange object of the table cell for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableCell object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTableCell implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline CComPtr<IRichTableCell> InitTableCell(ITextRange* pTableTextRange, ITextRange* pRowTextRange, ITextRange* pCellTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTableCell> pTableCell = NULL;
		InitTableCell(pTableTextRange, pRowTextRange, pCellTextRange, pRichTextBox, IID_IRichTableCell, reinterpret_cast<LPUNKNOWN*>(&pTableCell));
		return pTableCell;
	};

	/// \brief <em>Creates a \c TableCell object</em>
	///
	/// \overload
	///
	/// Creates a \c TableCell object that represents a given text range and returns its implementation of a
	/// given interface as a raw pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row for which to create the object.
	/// \param[in] pCellTextRange The \c ITextRange object of the table cell for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableCell object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTableCell Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline void InitTableCell(ITextRange* pTableTextRange, ITextRange* pRowTextRange, ITextRange* pCellTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppTableCell)
	{
		ATLASSERT_POINTER(ppTableCell, LPUNKNOWN);
		ATLASSUME(ppTableCell);

		*ppTableCell = NULL;
		if(pTableTextRange && pRowTextRange && pCellTextRange) {
			// create a TableCell object and initialize it with the given parameters
			CComObject<TableCell>* pTableCellObj = NULL;
			CComObject<TableCell>::CreateInstance(&pTableCellObj);
			pTableCellObj->AddRef();
			pTableCellObj->SetOwner(pRichTextBox);
			pTableCellObj->Attach(pTableTextRange, pRowTextRange, pCellTextRange);
			pTableCellObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTableCell));
			pTableCellObj->Release();
		}
	};

	/// \brief <em>Creates a \c TableCells object</em>
	///
	/// Creates a \c TableCells object that represents a given text range and returns its \c IRichTableCells
	/// implementation as a smart pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableCells object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichTableCells implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline CComPtr<IRichTableCells> InitTableCells(ITextRange* pTableTextRange, ITextRange* pRowTextRange, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichTableCells> pTableCells = NULL;
		InitTableCells(pTableTextRange, pRowTextRange, pRichTextBox, IID_IRichTableCells, reinterpret_cast<LPUNKNOWN*>(&pTableCells));
		return pTableCells;
	};

	/// \brief <em>Creates a \c TableCells object</em>
	///
	/// \overload
	///
	/// Creates a \c TableCells object that represents a given text range and returns its implementation of a
	/// given interface as a raw pointer.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table for which to create the object.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c TableCells object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTableCells Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	static inline void InitTableCells(ITextRange* pTableTextRange, ITextRange* pRowTextRange, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppTableCells)
	{
		ATLASSERT_POINTER(ppTableCells, LPUNKNOWN);
		ATLASSUME(ppTableCells);

		*ppTableCells = NULL;
		if(pTableTextRange && pRowTextRange) {
			// create a TableCells object and initialize it with the given parameters
			CComObject<TableCells>* pTableCellsObj = NULL;
			CComObject<TableCells>::CreateInstance(&pTableCellsObj);
			pTableCellsObj->AddRef();
			pTableCellsObj->SetOwner(pRichTextBox);
			pTableCellsObj->Attach(pTableTextRange, pRowTextRange);
			pTableCellsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTableCells));
			pTableCellsObj->Release();
		}
	};

	/// \brief <em>Creates an \c OLEObject object</em>
	///
	/// Creates an \c OLEObject object that represents a given embedded OLE object and returns its
	/// \c IRichOLEObject implementation as a smart pointer.
	///
	/// \param[in] pTextRange The \c ITextRange object of the OLE object for which to create the object.
	/// \param[in] pOleObject The \c IOleObject object for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c OLEObject object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichOLEObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">IOleObject</a>
	static inline CComPtr<IRichOLEObject> InitOLEObject(ITextRange* pTextRange, IOleObject* pOleObject, RichTextBox* pRichTextBox)
	{
		CComPtr<IRichOLEObject> pObject = NULL;
		InitOLEObject(pTextRange, pOleObject, pRichTextBox, IID_IRichOLEObject, reinterpret_cast<LPUNKNOWN*>(&pObject));
		return pObject;
	};

	/// \brief <em>Creates an \c OLEObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEObject object that represents a given embedded OLE object and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] pTextRange The \c ITextRange object of the OLE object for which to create the object.
	/// \param[in] pOleObject The \c IOleObject object for which to create the object.
	/// \param[in] pRichTextBox The \c RichTextBox object the \c OLEObject object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">IOleObject</a>
	static inline void InitOLEObject(ITextRange* pTextRange, IOleObject* pOleObject, RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppObject)
	{
		ATLASSERT_POINTER(ppObject, LPUNKNOWN);
		ATLASSUME(ppObject);

		*ppObject = NULL;
		if(pOleObject) {
			// create a OLEObject object and initialize it with the given parameters
			CComObject<OLEObject>* pOLEObjectObj = NULL;
			CComObject<OLEObject>::CreateInstance(&pOLEObjectObj);
			pOLEObjectObj->AddRef();
			pOLEObjectObj->Attach(pTextRange, pOleObject);
			pOLEObjectObj->SetOwner(pRichTextBox);
			pOLEObjectObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppObject));
			pOLEObjectObj->Release();
		}
	};

	/// \brief <em>Creates a \c OLEObjects object</em>
	///
	/// Creates a \c OLEObjects object that represents all embedded OLE objects of the given control and
	/// returns its \c IRichOLEObjects implementation as a smart pointer.
	///
	/// \param[in] pRichTextBox The \c RichTextBox object the \c OLEObject object will work on.
	///
	/// \return A smart pointer to the created object's \c IRichOLEObjects implementation.
	static inline CComPtr<IRichOLEObjects> InitOLEObjects(RichTextBox* pRichTextBox)
	{
		CComPtr<IRichOLEObjects> pOLEObjects = NULL;
		InitOLEObjects(pRichTextBox, IID_IRichOLEObjects, reinterpret_cast<LPUNKNOWN*>(&pOLEObjects));
		return pOLEObjects;
	};

	/// \brief <em>Creates a \c OLEObjects object</em>
	///
	/// \overload
	///
	/// Creates a \c OLEObjects object that represents all embedded OLE objects of the given control and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pRichTextBox The \c RichTextBox object the \c OLEObjects object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppOLEObjects Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitOLEObjects(RichTextBox* pRichTextBox, REFIID requiredInterface, LPUNKNOWN* ppOLEObjects)
	{
		ATLASSERT_POINTER(ppOLEObjects, LPUNKNOWN);
		ATLASSUME(ppOLEObjects);

		*ppOLEObjects = NULL;
		if(pRichTextBox) {
			// create a TableRows object and initialize it with the given parameters
			CComObject<OLEObjects>* pOLEObjectsObj = NULL;
			CComObject<OLEObjects>::CreateInstance(&pOLEObjectsObj);
			pOLEObjectsObj->AddRef();
			pOLEObjectsObj->SetOwner(pRichTextBox);
			pOLEObjectsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppOLEObjects));
			pOLEObjectsObj->Release();
		}
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};
};     // ClassFactory