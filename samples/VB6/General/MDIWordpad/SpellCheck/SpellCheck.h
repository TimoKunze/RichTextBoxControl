// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the SPELLCHECK_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// SPELLCHECK_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef SPELLCHECK_EXPORTS
	#define SPELLCHECK_API __stdcall //__declspec(dllexport)
#else
	#define SPELLCHECK_API __declspec(dllimport)
#endif


extern std::hash_map<LANGID, HyphenDict*> g_hunSpellHyphenationDictionaries;

BOOL SPELLCHECK_API HunSpellLoadHyphenationDictionary(_In_ LANGID langid, _In_ WCHAR* pFile);
BOOL SPELLCHECK_API HunSpellFreeHyphenationDictionary(_In_ LANGID langid);
void SPELLCHECK_API HunSpellHyphenate(_In_ WCHAR* pszWord, _In_ LANGID langid, _In_ LONG ichExceed, _Out_ HYPHRESULT* phyphresult);