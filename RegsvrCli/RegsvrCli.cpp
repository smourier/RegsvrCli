#include <windows.h>
#include <string>
#include <strsafe.h>
#include <stdio.h>

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

int main()
{
	auto reg = true;
	wchar_t path[2048] = L"";

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

	HRESULT hr = S_OK;
	auto h = LoadLibrary(path);
	if (!h)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		wprintf(L"Calling LoadLibrary(\"%s\") returned error 0x%08X (%i)%s\n", path, hr, hr, GetErrorMessage(hr).c_str());
		return hr;
	}

	typedef HRESULT(WINAPI* DllRegUnRegServer)();
	auto fn = (DllRegUnRegServer)GetProcAddress(h, reg ? "DllRegisterServer" : "DllUnregisterServer");
	if (!fn)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		wprintf(L"Calling GetProcAddress(\"%s\") returned error 0x%08X (%i)%s\n", reg ? L"DllRegisterServer" : L"DllUnregisterServer", hr, hr, GetErrorMessage(hr).c_str());
	}
	else
	{
		hr = fn();
		wprintf(L"%s '%s' returned: 0x%08X (%i)\n", reg ? L"Registering" : L"Unregistering", path, hr, hr);
	}

	FreeLibrary(h);
	return hr;
}
