# RegsvrCli
A console tool to call DllRegisterServer/DllUnRegisterServer and display the error code.

Usage is `RegsvrCli <path> [/u]`

By default, RegsvrCli calls `DllRegisterServer` on <path> file. With `/u` option, RegsvrCli calls `DllUnregisterServer`.
