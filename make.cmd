@echo off
REM %1 is visual studio version (2015, 2017 or 2019)
REM %2 is configuration (Release or Debug)
REM %3 is platform (x86 or x64)

REM Validating input
IF "%~1" == "2015" (
    SET VISUAL_STUDIO="Visual Studio 14 2015"
) ELSE (
    IF "%~1" == "2017" (
        SET VISUAL_STUDIO="Visual Studio 15 2017"
    ) ELSE (
        IF "%~1" == "2019" (
            SET VISUAL_STUDIO="Visual Studio 16 2019"
        ) ELSE (
            IF "%~1" == "2022" (
                SET VISUAL_STUDIO="Visual Studio 17 2022"
            ) ELSE (
                GOTO Error
            )   
        )   
    )
)

IF "%~2" NEQ "Release" (
    IF "%~2" NEQ "Debug" (
        GOTO Error
    )
)

IF "%~3" NEQ "x64" (
    IF "%~3" NEQ "Win32" (
        GOTO Error
    )
)

cmake -S src -B build/%1%3 -G %VISUAL_STUDIO% -A %3
if %errorlevel% NEQ 0 exit /b 1

msbuild build/%1%3/excelbind.sln /t:Rebuild /p:Configuration=%2 /p:Platform=%3
if %errorlevel% NEQ 0 exit /b 1

GOTO EOF

:Error
echo "Calling convention is 'make compilerVersion configuration platform'"
echo "Where compilerVersion in [2015, 2017, 2019, 2022], configuration in [Release, Debug] and platform in [x86, x64]"
exit /b 1
:EOF