# Copyright (c) Meta Platforms, Inc. and affiliates.
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.

macro(torch_hip_get_arch_list store_var)
  if(DEFINED ENV{PYTORCH_ROCM_ARCH})
    set(_TMP $ENV{PYTORCH_ROCM_ARCH})
  else()
    # Use arch of installed GPUs as default
    execute_process(COMMAND "rocm_agent_enumerator" COMMAND bash "-c" "grep -v gfx000 | sort -u | xargs | tr -d '\n'"
                    RESULT_VARIABLE ROCM_AGENT_ENUMERATOR_RESULT
                    OUTPUT_VARIABLE ROCM_ARCH_INSTALLED)
    if(NOT ROCM_AGENT_ENUMERATOR_RESULT EQUAL 0)
      message(FATAL_ERROR " Could not detect ROCm arch for GPUs on machine. Result: '${ROCM_AGENT_ENUMERATOR_RESULT}'")
    endif()
    set(_TMP ${ROCM_ARCH_INSTALLED})
  endif()
  string(REPLACE " " ";" ${store_var} "${_TMP}")
endmacro()

macro(GET_AMD_WAVESIZE AMD_WAVESIZE)
  list(APPEND AMD_NAVI_ARCH_LIST "gfx1010" "gfx1011" "gfx1012" "gfx1030" 
                                 "gfx1031" "gfx1032" "gfx1100" "gfx1101" "gfx1102")
  include(${CMAKE_CURRENT_SOURCE_DIR}/cmake/ArchMacros.cmake)
  torch_hip_get_arch_list(ROCM_ARCH)
  if(ROCM_ARCH IN_LIST AMD_NAVI_ARCH_LIST)
    set(${AMD_WAVESIZE} 32 CACHE INTERNAL "")
  else()
    set(${AMD_WAVESIZE} 64 CACHE INTERNAL "")
  endif()
endmacro()
