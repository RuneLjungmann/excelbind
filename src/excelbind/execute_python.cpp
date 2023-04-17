#include <corecrt.h>
#include "pybind11/embed.h"

#include "xll12/xll/xll.h"

#include "type_conversion.h"
#include "configuration.h"
/*
This piece of code registers an Excel function called 'execute_python', that makes it possible to run python code directly from Excel.
The approach take, is that the provided code is wrapped internally as a function. I.e. the Excel call
=execute_python("return arg0 + arg1", A1, A2)
will create the script:
    def func(arg0, arg1):
        return arg0 + arg1

    out = func(arg0_in, arg1_in)

where the value of cells A1 and A2 are passed to the local variables arg0_in and arg1_in.
The local variable out is returned as the result of the call to the execute_python function.
*/

namespace py = pybind11;
using namespace pybind11::literals;

// unique str used to avoid name collisions when building internal python code
#define SOME_UNIQUE_STR "89C86679C2D042AD966268A41890F4F1"




py::str convert_script_to_function(const XCHAR* script)
{
    py::str script_py = cast_string(std::wstring(script));
    py::list script_split = script_py.attr("splitlines")();
    py::list script_wrapped;
    script_wrapped.append("def func_" SOME_UNIQUE_STR "(arg0, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9):");
    for (auto& i : script_split)
    {
        script_wrapped.append("    " + i.cast<std::string>());
    }
    script_wrapped.append("out_" SOME_UNIQUE_STR " = func_" SOME_UNIQUE_STR
        "(arg0_" SOME_UNIQUE_STR
        ", arg1_" SOME_UNIQUE_STR
        ", arg2_" SOME_UNIQUE_STR
        ", arg3_" SOME_UNIQUE_STR
        ", arg4_" SOME_UNIQUE_STR
        ", arg5_" SOME_UNIQUE_STR
        ", arg6_" SOME_UNIQUE_STR
        ", arg7_" SOME_UNIQUE_STR
        ", arg8_" SOME_UNIQUE_STR
        ", arg9_" SOME_UNIQUE_STR ")");
    auto os = py::module::import("os");
    py::str all = os.attr("linesep");
    return all.attr("join")(script_wrapped);
}






#pragma warning(push)
#pragma warning(disable: 4100)
xll::LPOPER executePython(
    const XCHAR* script, 
    xll::OPER* arg0,
    xll::OPER* arg1,
    xll::OPER* arg2,
    xll::OPER* arg3,
    xll::OPER* arg4,
    xll::OPER* arg5,
    xll::OPER* arg6,
    xll::OPER* arg7,
    xll::OPER* arg8,
    xll::OPER* arg9
)
{
#pragma XLLEXPORT
    static xll::OPER res_xll;
    try {
        py::dict locals;        
        locals[py::str("arg0_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg0);
        locals[py::str("arg1_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg1);
        locals[py::str("arg2_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg2);
        locals[py::str("arg3_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg3);
        locals[py::str("arg4_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg4);
        locals[py::str("arg5_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg5);
        locals[py::str("arg6_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg6);
        locals[py::str("arg7_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg7);
        locals[py::str("arg8_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg8);
        locals[py::str("arg9_" SOME_UNIQUE_STR)] = cast_oper_to_py(*arg9);

        py::str script_py = convert_script_to_function(script);
        py::exec(script_py, py::globals(), locals);
        py::object res_py = locals["out_" SOME_UNIQUE_STR];
        cast_py_to_oper(res_py, res_xll);
    }
    catch (const std::exception& e)
    {
        if (Configuration::is_error_messages_enabled())
        {
            XLL_ERROR(e.what());
        }
        res_xll = xll::OPER(xlerr::Num);
    }
    return &res_xll;
}
#pragma warning(pop)


xll::AddIn execute_python(
    xll::Function(XLL_LPOPER, L"?executePython", 
        Configuration::is_function_prefix_set()
        ? (cast_string(Configuration::function_prefix()) + L".execute_python").c_str() : L"execute_python")
    .Arg(XLL_CSTRING, L"Script")
    .Arg(XLL_LPOPER, L"arg0")
    .Arg(XLL_LPOPER, L"arg1")
    .Arg(XLL_LPOPER, L"arg2")
    .Arg(XLL_LPOPER, L"arg3")
    .Arg(XLL_LPOPER, L"arg4")
    .Arg(XLL_LPOPER, L"arg5")
    .Arg(XLL_LPOPER, L"arg6")
    .Arg(XLL_LPOPER, L"arg7")
    .Arg(XLL_LPOPER, L"arg8")
    .Arg(XLL_LPOPER, L"arg9")
    .Category(cast_string(Configuration::excel_category()).c_str())
    .FunctionHelp(L"Execute python code")
);
