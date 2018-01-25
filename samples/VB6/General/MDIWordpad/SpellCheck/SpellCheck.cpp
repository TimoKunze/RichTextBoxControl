// SpellCheck.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "SpellCheck.h"


std::hash_map<LANGID, HyphenDict*> g_hunSpellHyphenationDictionaries;

BOOL SPELLCHECK_API HunSpellLoadHyphenationDictionary(_In_ LANGID langid, _In_ WCHAR* pFile)
{
	std::hash_map<LANGID, HyphenDict*>::iterator it = g_hunSpellHyphenationDictionaries.find(langid);
	if(it != g_hunSpellHyphenationDictionaries.end()) {
		hnj_hyphen_free(it->second);
		g_hunSpellHyphenationDictionaries.erase(it);
	}

	BOOL ret = FALSE;

	int requiredSize = WideCharToMultiByte(CP_ACP, 0, pFile, -1, NULL, 0, NULL, NULL);
	if(requiredSize > 0) {
		LPSTR pAnsiFile = static_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, requiredSize * sizeof(CHAR)));
		if(pAnsiFile) {
			if(WideCharToMultiByte(CP_ACP, 0, pFile, -1, pAnsiFile, requiredSize * sizeof(CHAR), NULL, NULL)) {
				HyphenDict* pDictionary = hnj_hyphen_load(pAnsiFile);
				if(pDictionary) {
					g_hunSpellHyphenationDictionaries[langid] = pDictionary;
					ret = TRUE;
				}
			}
			HeapFree(GetProcessHeap(), 0, pAnsiFile);
		}
	}
	return ret;
}

BOOL SPELLCHECK_API HunSpellFreeHyphenationDictionary(_In_ LANGID langid)
{
	std::hash_map<LANGID, HyphenDict*>::iterator it = g_hunSpellHyphenationDictionaries.find(langid);
	if(it != g_hunSpellHyphenationDictionaries.end()) {
		hnj_hyphen_free(it->second);
		g_hunSpellHyphenationDictionaries.erase(it);
		return TRUE;
	}
	return FALSE;
}

void SPELLCHECK_API HunSpellHyphenate(_In_ WCHAR* pszWord, _In_ LANGID langid, _In_ LONG ichExceed, _Out_ HYPHRESULT* phyphresult)
{
	phyphresult->khyph = khyphNil;

	std::hash_map<LANGID, HyphenDict*>::iterator it = g_hunSpellHyphenationDictionaries.find(langid);
	if(it != g_hunSpellHyphenationDictionaries.end()) {
		HyphenDict* pDictionary = it->second;
		if(pDictionary) {
			int requiredSize = WideCharToMultiByte(CP_ACP, 0, pszWord, -1, NULL, 0, NULL, NULL);
			if(requiredSize > 0) {
				LPSTR pAnsiWord = static_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, requiredSize * sizeof(CHAR)));
				if(pAnsiWord) {
					if(WideCharToMultiByte(CP_ACP, 0, pszWord, -1, pAnsiWord, requiredSize * sizeof(CHAR), NULL, NULL)) {
						for(int i = 0; i < requiredSize - 1; i++) {
							pAnsiWord[i] = static_cast<CHAR>(tolower(pAnsiWord[i]));
						}
						LPSTR pHyphens = static_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, (requiredSize + 10) * sizeof(CHAR)));
						if(pHyphens) {
							LPSTR pHyphenatedWord = static_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, (requiredSize * 2) * sizeof(CHAR)));
							if(pHyphenatedWord) {
								char** rep = NULL;
								int* pos = NULL;
								int* cut = NULL;
								pHyphenatedWord[0] = '\0';
								if(!hnj_hyphen_hyphenate2(pDictionary, pAnsiWord, requiredSize - 1, pHyphens, pHyphenatedWord, &rep, &pos, &cut)) {
									int j = 0;
									for(int i = 0; i < min(ichExceed, requiredSize); i++) {
										if(pHyphens[i] & 1) {
											if(pHyphenatedWord[j] != pAnsiWord[i] && pHyphenatedWord[j + 1] == '=' && pHyphenatedWord[j + 2] != pAnsiWord[i + 1]) {
												// for instance omaatje -> omi-tje (constructed sample)
												j++;
												phyphresult->khyph = khyphDelAndChange;
												phyphresult->ichHyph = i + 1;
												phyphresult->chHyph = pHyphenatedWord[j - 1];
												pHyphenatedWord[j++] = '-';
												j++;
												i += 2;
											} else if(pHyphenatedWord[j] != pAnsiWord[i] && pHyphenatedWord[j + 1] == '=') {
												// for instance Zucker -> Zuk-ker
												j++;
												phyphresult->khyph = khyphChangeBefore;
												phyphresult->ichHyph = i;
												phyphresult->chHyph = pHyphenatedWord[j - 1];
												pHyphenatedWord[j++] = '-';
												j++;
												i++;
											} else if(pHyphenatedWord[j + 1] == '=') {
												// for instance Zucker -> Zu-cker
												pHyphenatedWord[j++] = pAnsiWord[i];
												pHyphenatedWord[j++] = '-';
												phyphresult->khyph = khyphNormal;
												phyphresult->ichHyph = i;
												if(pHyphenatedWord[j] != pAnsiWord[i + 1]) {
													if(pHyphenatedWord[j + 1] == pAnsiWord[i + 2]) {
														// for instance Zucker -> Zuc-ler (constructed sample)
														phyphresult->khyph = khyphChangeAfter;
														phyphresult->ichHyph = i;
														phyphresult->chHyph = pHyphenatedWord[j];
														j++;
														i++;
													} else {
														// for instance omaatje -> oma-tje
														phyphresult->khyph = khyphDeleteBefore;			// actually it is delete after, but this does not exist
														phyphresult->ichHyph = i + 1;
														j++;
														i += 2;
													}
												}
											} else {
												// for instance Schiffahrt -> Schiff-fahrt
												pHyphenatedWord[j++] = pAnsiWord[i];
												phyphresult->khyph = khyphAddBefore;
												phyphresult->ichHyph = i;
												phyphresult->chHyph = pHyphenatedWord[j];
												j++;
												pHyphenatedWord[j++] = '-';
											}
										} else {
											pHyphenatedWord[j++] = pAnsiWord[i];
										}
									}
								}
								HeapFree(GetProcessHeap(), 0, pHyphenatedWord);
							}
							HeapFree(GetProcessHeap(), 0, pHyphens);
						}
					}
					HeapFree(GetProcessHeap(), 0, pAnsiWord);
				}
			}
		}
	}
}