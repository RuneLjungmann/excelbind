cmake_minimum_required (VERSION 3.1)

project(excelbind CXX ASM_MASM)

# set env var PYTHON_VERSION or change the lines below if you plan to build against python35 or python36.
if(NOT PYTHON_VERSION)
    set(PYTHON_VERSION python37)
endif()

# Checks if the USEPYTHONVERSION_PYTHONLOCATION env var is set
# (by e.g. Azure pipeline) otherwise default to standard location for python installation.
if(${USEPYTHONVERSION_PYTHONLOCATION})
    set(PYTHON_HOME ${USEPYTHONVERSION_PYTHONLOCATION})
elseif("${CMAKE_GENERATOR_PLATFORM}" STREQUAL "Win32")
    set(PYTHON_HOME c:/${PYTHON_VERSION})
else()
    set(PYTHON_HOME c:/${PYTHON_VERSION}_x64)    
endif()

file(GLOB SOURCES
    ${CMAKE_SOURCE_DIR}/excelbind/chrono.h
    ${CMAKE_SOURCE_DIR}/excelbind/configuration.h
    ${CMAKE_SOURCE_DIR}/excelbind/python_function_adapter.h
    ${CMAKE_SOURCE_DIR}/excelbind/script_manager.h
    ${CMAKE_SOURCE_DIR}/excelbind/type_conversion.h

    ${CMAKE_SOURCE_DIR}/excelbind/configuration.cpp
    ${CMAKE_SOURCE_DIR}/excelbind/execute_python.cpp
    ${CMAKE_SOURCE_DIR}/excelbind/function_registration.cpp
    ${CMAKE_SOURCE_DIR}/excelbind/python_function_adapter.cpp
    ${CMAKE_SOURCE_DIR}/excelbind/script_manager.cpp
    ${CMAKE_SOURCE_DIR}/excelbind/type_conversion.cpp

    ${CMAKE_SOURCE_DIR}/excelbind/fct_exports.asm    
)

include_directories(${CMAKE_SOURCE_DIR}/thirdparty)
include_directories(${CMAKE_SOURCE_DIR}/thirdparty/pybind11/include)
include_directories(${PYTHON_HOME}/include)

add_library(excelbind SHARED ${SOURCES})
    
add_compile_definitions(_LIB)
add_compile_definitions($<$<CONFIG:Debug>:_DEBUG>)
target_compile_options(excelbind PUBLIC -WX -W4)

add_subdirectory(${CMAKE_SOURCE_DIR}/thirdparty/xll12)


target_link_options(excelbind PRIVATE /SAFESEH:NO /DEBUG:FULL)

#Pybind needs to link to both python3x.lib and python3x_d.lib in debug mode!
target_link_libraries(excelbind ${PYTHON_HOME}/libs/${PYTHON_VERSION}.lib)
target_link_libraries(excelbind ${PYTHON_HOME}/libs/${PYTHON_VERSION}$<$<CONFIG:Debug>:_d>.lib)

target_link_libraries(excelbind xll12)

add_custom_command(TARGET excelbind POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:excelbind> ${CMAKE_BINARY_DIR}/../../Binaries/${CMAKE_GENERATOR_PLATFORM}/$<CONFIG>/excelbind.xll
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE_DIR:excelbind>/excelbind.pdb ${CMAKE_BINARY_DIR}/../../Binaries/${CMAKE_GENERATOR_PLATFORM}/$<CONFIG>/excelbind.pdb
)