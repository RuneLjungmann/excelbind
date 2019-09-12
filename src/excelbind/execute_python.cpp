#include <Python.h>
#include "pybind11/embed.h"

#include "xll12/xll/xll.h"

#include "type_conversion.h"
#include "configuration.h"

namespace py = pybind11;
using namespace pybind11::literals;

// unique str used to avoid name collisions when building internal python code
#define SOME_UNIQUE_STR "89C86679C2D042AD966268A41890F4F1"


std::wstring to_string(const xll::OPER& oper)
{
    return std::wstring(oper.val.str + 1, oper.val.str[0]);
}


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


void convert_output(py::object in, xll::OPER& out)
{
    if (py::isinstance<py::str>(in))
    {
        out = in.cast<std::wstring>().c_str();
    }
    else if (py::isinstance<py::int_>(in))
    {
        out = in.cast<int>();
    }
    else if (py::isinstance<py::float_>(in))
    {
        out = in.cast<double>();
    }
}


py::object convert_input(const xll::OPER & in)
{
    if (in.isBool())
    {
        return py::bool_(static_cast<bool>(in));
    }
    else if (in.isInt())
    {
        return py::int_(static_cast<int>(in));
    }
    else if (in.isNum())
    {
        return py::float_(static_cast<double>(in));
    }
    else if (in.isStr())
    {
        return py::str(cast_string(to_string(in)));
    }

    return py::none();
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
        locals[py::str("arg0_" SOME_UNIQUE_STR)] = convert_input(*arg0);
        locals[py::str("arg1_" SOME_UNIQUE_STR)] = convert_input(*arg1);
        locals[py::str("arg2_" SOME_UNIQUE_STR)] = convert_input(*arg2);
        locals[py::str("arg3_" SOME_UNIQUE_STR)] = convert_input(*arg3);
        locals[py::str("arg4_" SOME_UNIQUE_STR)] = convert_input(*arg4);
        locals[py::str("arg5_" SOME_UNIQUE_STR)] = convert_input(*arg5);
        locals[py::str("arg6_" SOME_UNIQUE_STR)] = convert_input(*arg6);
        locals[py::str("arg7_" SOME_UNIQUE_STR)] = convert_input(*arg7);
        locals[py::str("arg8_" SOME_UNIQUE_STR)] = convert_input(*arg8);
        locals[py::str("arg9_" SOME_UNIQUE_STR)] = convert_input(*arg9);

        py::str script_py = convert_script_to_function(script);
        py::exec(script_py, py::globals(), locals);
        py::object res_py = locals["out_" SOME_UNIQUE_STR];
        convert_output(res_py, res_xll);
    }
    catch (std::exception& e)
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
