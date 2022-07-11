cmake_minimum_required(VERSION 3.20)

project(excel)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

include("${CMAKE_TOOLCHAIN_FILE}")

# Setup boost
set(Boost_USE_STATIC_LIBS on)
find_package(
	Boost 1.65 REQUIRED
	COMPONENTS system
)

find_package(
	libzip 
	CONFIG
	REQUIRED
)

configure_file("${CMAKE_SOURCE_DIR}/include/DirectoryConfig.hpp.in" "${CMAKE_SOURCE_DIR}/include/DirectoryConfig.hpp")

option(TESTING "Whether to compile the tests" OFF)

if(TESTING)
	# Test Definition
	set(includes_list ${Boost_INCLUDE_DIRS} "${CMAKE_SOURCE_DIR}/include")
	set(SOURCES "test/test.cpp")
	add_executable(test ${SOURCES})
	target_include_directories(test PUBLIC ${includes_list})
	target_link_libraries(test ${Boost_LIBRARIES} libzip::zip)
endif()