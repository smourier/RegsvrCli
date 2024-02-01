#include <windows.h>
#include <string>
#include <strsafe.h>
#include <stdio.h>

typedef HRESULT(WINAPI* DllRegUnRegServer)();

void Help()
{
	wprintf(L"Usage is RegsvrCli <path> [/u]\n");
	wprintf(L"\n");
	wprintf(L"By default, RegsvrCli calls DllRegisterServer on <path> file. With /u option, RegsvrCli calls DllUnregisterServer.\n");
}

std::wstring GetErrorMessage(int error)
{
	wchar_t message[1024];
	if (!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_ARGUMENT_ARRAY, nullptr, error, 0, message, _countof(message), nullptr))
		return L"";

	return L": " + std::wstring(message);
}

std::wstring to_wstring(const std::string& s)
{
	if (s.empty())
		return std::wstring();

	auto ssize = (int)s.size();
	auto wsize = MultiByteToWideChar(CP_THREAD_ACP, 0, s.data(), ssize, nullptr, 0);
	if (!wsize)
		return std::wstring();

	std::wstring ws;
	ws.resize(wsize);
	wsize = MultiByteToWideChar(CP_THREAD_ACP, 0, s.data(), ssize, &ws[0], wsize);
	if (!wsize)
		return std::wstring();

	return ws;
}

int main()
{
	SYSTEMTIME st{};
	GetLocalTime(&st);
	auto bitness = (sizeof(void*) == 8) ? 64 : 32;
	wprintf(L"RegsvrCli %u-bit - Copyright (C) 2023-%u Simon Mourier. All rights reserved.\n\n", bitness, st.wYear);

	int argc;
	auto argv = CommandLineToArgvW(GetCommandLine(), &argc);
	if (argc < 1)
	{
		Help();
		return -1;
	}

	auto reg = true;
	wchar_t path[2048] = L"";
	for (auto i = 1; i < argc; i++)
	{
		if (!_wcsicmp(argv[i], L"/u"))
		{
			reg = false;
			continue;
		}

		StringCchCopy(path, _countof(path), argv[i]);
	}

	if (!*path)
	{
		Help();
		return -1;
	}

	auto coInit = false;
	auto oleInit = true;
	auto hr = OleInitialize(nullptr);
	if (FAILED(hr))
	{
		wprintf(L"(warning) Calling OleInitialize() returned error 0x%08X (%i)%s\n", hr, hr, GetErrorMessage(hr).c_str());
		oleInit = false;
		hr = CoInitialize(nullptr);
		if (FAILED(hr))
		{
			wprintf(L"(warning) Calling CoInitialize() returned error 0x%08X (%i)%s\n", hr, hr, GetErrorMessage(hr).c_str());
			hr = S_OK;
		}
		else
		{
			coInit = true;
		}
	}

	auto h = LoadLibraryEx(path, nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
	if (!h)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		wprintf(L"Calling LoadLibrary(\"%s\") returned error 0x%08X (%i)%s\n", path, hr, hr, GetErrorMessage(hr).c_str());
	}
	else
	{
		char name[64] = "";
		DllRegUnRegServer fn;
		if (reg)
		{
			StringCchCopyA(name, _countof(name), "DllRegisterServer");
			fn = (DllRegUnRegServer)GetProcAddress(h, name);
		}
		else
		{
			StringCchCopyA(name, _countof(name), "DllUnregisterServer");
			fn = (DllRegUnRegServer)GetProcAddress(h, name);
		}

		if (!fn)
		{
			hr = HRESULT_FROM_WIN32(GetLastError());
			wprintf(L"Calling GetProcAddress(\"%s\") returned error 0x%08X (%i)%s\n", to_wstring(name).c_str(), hr, hr, GetErrorMessage(hr).c_str());
		}
		else
		{
			hr = fn();
			wprintf(L"Calling %s() on '%s' returned: 0x%08X (%i)\n", to_wstring(name).c_str(), path, hr, hr);
		}

		FreeLibrary(h);
	}

	if (coInit)
	{
		CoUninitialize();
	}
	else if (oleInit)
	{
		OleUninitialize();
	}
	return hr;
}
