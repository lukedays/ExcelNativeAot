#define PathToLibrary "..\\Addin\\bin\\Debug\\net8.0\\win-x64\\native\\Addin.dll"

#include "windows.h"
#pragma comment (lib, "ole32.lib")

#include <stdlib.h>
#include <stdio.h>

#ifndef F_OK
#define F_OK    0
#endif

int callLibFunction(char* path, char* funcName)
{
	HINSTANCE handle = LoadLibraryA(path);

	typedef void(*myFunc)();
	myFunc MyImport = (myFunc)GetProcAddress(handle, funcName);

	MyImport();

	//return result;
}

int main()
{
	if (access(PathToLibrary, F_OK) == -1)
	{
		puts("Couldn't find library at the specified path");
		return 0;
	}

    printf("%d", callLibFunction(PathToLibrary, "ComTest"));
}
