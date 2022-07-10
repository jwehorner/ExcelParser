include(FetchContent)

FetchContent_Declare(
  libzip
  URL https://codeload.github.com/nih-at/libzip/zip/refs/heads/main
)

FetchContent_GetProperties(libzip)
if(NOT libzip_POPULATED)
  FetchContent_Populate(libzip)
endif()

set(LIBZIP_INCLUDE_DIR ${FETCHCONTENT_BASE_DIR}/libzip-src/lib CACHE STRING "libzip Include File Location")
set(LIBZIP_BUILD_DIR ${LIBZIP_INCLUDE_DIR}/lib/Debug CACHE STRING "libzip Build Location")

add_subdirectory(${FETCHCONTENT_BASE_DIR}/libzip-src ${LIBZIP_INCLUDE_DIR})

