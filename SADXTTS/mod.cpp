#include "pch.h"
#include "SADXModLoader.h"
#include "FunctionHook.h"
#include "IniFile.hpp"
#include <sapi.h>
#include <sphelper.h>
#include <atlbase.h>
#include <vector>

CComPtr<ISpVoice> pVoice = nullptr;
CComPtr<ISpObjectToken> voicetoken = nullptr;
SPVOICESTATUS voicestatus;
std::vector<wchar_t> textbuf;
FunctionHook<void, MSGC*, const char*> MSG_Puts_h(MSG_Puts);
void MSG_Puts_r(MSGC* msgc, const char* str)
{
	MSG_Puts_h.Original(msgc, str);
	int cp = TextLanguage > Languages_English ? 1252 : 932;
	auto len2 = MultiByteToWideChar(cp, 0, str, -1, nullptr, 0);
	textbuf.resize(len2, 0);
	MultiByteToWideChar(cp, 0, str, -1, textbuf.data(), textbuf.size());
	for (wchar_t& ch : textbuf)
		if (ch == '\n')
			ch = ' ';
	pVoice->Speak(textbuf.data(), SPF_ASYNC | SPF_IS_NOT_XML, nullptr);
}

FunctionHook<void> EV_SerifStop_h(EV_SerifStop);
void EV_SerifStop_r()
{
	pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr);
}

FunctionHook<void> EV_SerifWait_h(EV_SerifWait);
void EV_SerifWait_r()
{
	while (true)
	{
#pragma warning(suppress : 6387)
		pVoice->GetStatus(&voicestatus, nullptr);
		if (voicestatus.dwRunningState != SPRUNSTATE::SPRS_IS_SPEAKING)
			break;
		EV_Wait(1);
	}
}

FunctionHook<void, int> EV_SerifPlay_h(EV_SerifPlay);
void EV_SerifPlay_r(int id)
{
	EV_SerifWait_r();
}

FunctionHook<int, int, int*> GetHintText_h(GetHintText);
int GetHintText_r(int id, int* data)
{
	int result = GetHintText_h.Original(id, data);
	data[1] = -1;
	return result;
}

extern "C"
{
	__declspec(dllexport) void Init(const char* path, const HelperFunctions& helperFunctions)
	{
		if (SUCCEEDED(pVoice.CoCreateInstance(CLSID_SpVoice)))
		{
			char pathbuf[MAX_PATH];
			sprintf_s(pathbuf, "%s\\config.ini", path);
			IniFile config(pathbuf);
			const auto group = config.getGroup("Voice");
			if (group->cbegin() != group->cend())
			{
				std::wstring attribs;
				if (group->hasKeyNonEmpty("Name"))
					attribs += L"Name=" + group->getWString("Name");
				if (group->hasKeyNonEmpty("Vendor"))
				{
					if (!attribs.empty())
						attribs += L';';
					attribs += L"Vendor=" + group->getWString("Vendor");
				}
				if (group->hasKeyNonEmpty("Language"))
				{
					if (!attribs.empty())
						attribs += L';';
					attribs += L"Language=" + group->getWString("Language");
				}
				if (group->hasKeyNonEmpty("Age"))
				{
					if (!attribs.empty())
						attribs += L';';
					attribs += L"Age=" + group->getWString("Age");
				}
				if (group->hasKeyNonEmpty("Gender"))
				{
					if (!attribs.empty())
						attribs += L';';
					attribs += L"Gender=" + group->getWString("Gender");
				}
				if (SUCCEEDED(SpFindBestToken(SPCAT_VOICES, attribs.c_str(), nullptr, &voicetoken)))
					pVoice->SetVoice(voicetoken);
				else
					PrintDebug("SADX TTS: Failed to find voice matching criteria!");
			}
			MSG_Puts_h.Hook(MSG_Puts_r);
			EV_SerifStop_h.Hook(EV_SerifStop_r);
			EV_SerifWait_h.Hook(EV_SerifWait_r);
			EV_SerifPlay_h.Hook(EV_SerifPlay_r);
			GetHintText_h.Hook(GetHintText_r);
		}
		else
			PrintDebug("SADX TTS: Failed to initialize text-to-speech engine!");
	}

	__declspec(dllexport) ModInfo SADXModInfo = { ModLoaderVer };
}