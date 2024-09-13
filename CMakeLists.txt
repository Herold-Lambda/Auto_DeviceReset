cmake_minimum_required(VERSION 3.28)
project(AudioDeviceResetGUI)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

# Set Qt path for MSVC
set(QT_PATH "C:/Qt/6.7.2/msvc2019_64")
set(CMAKE_PREFIX_PATH ${QT_PATH})

find_package(Qt6 COMPONENTS Core Gui Widgets REQUIRED)

set(SOURCES
        main.cpp
        resources.qrc
)

add_executable(AudioDeviceResetGUI WIN32 ${SOURCES})

target_link_libraries(AudioDeviceResetGUI PRIVATE
        Qt6::Core
        Qt6::Gui
        Qt6::Widgets
        ole32
        winmm
        oleaut32
        uuid
        Mmdevapi
)

# Use windeployqt for deployment
if(WIN32)
    get_target_property(_qmake_executable Qt6::qmake IMPORTED_LOCATION)
    get_filename_component(_qt_bin_dir "${_qmake_executable}" DIRECTORY)
    find_program(WINDEPLOYQT_EXECUTABLE windeployqt HINTS "${_qt_bin_dir}")

    add_custom_command(TARGET AudioDeviceResetGUI POST_BUILD
            COMMAND "${CMAKE_COMMAND}" -E env PATH="${_qt_bin_dir}" "${WINDEPLOYQT_EXECUTABLE}"
            --verbose 0
            --no-compiler-runtime
            \"$<TARGET_FILE:AudioDeviceResetGUI>\"
            COMMENT "Running windeployqt..."
    )
endif()

# Copy icon file to build directory
add_custom_command(TARGET AudioDeviceResetGUI POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy
        ${CMAKE_SOURCE_DIR}/icon.ico
        $<TARGET_FILE_DIR:AudioDeviceResetGUI>/icon.ico
)