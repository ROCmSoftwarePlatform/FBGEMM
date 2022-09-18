/*******************************************************************************
 * Copyright (c) 2016 - 2022 Advanced Micro Devices, Inc. All rights reserved.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 ******************************************************************************/

#include <hip/hip_runtime.h>
#include <hip/hip_fp16.h>
#include "split_tbe_common_hip.h"

template <typename cache_t, typename emb_t, int32_t embedding_dim, int32_t weigiht_decay_mode>
struct rowwise_adagrad_optimizer_t
{
    __device__ rowwise_adagrad_optimizer_t(const rowwise_adagrad_kernel_arg_t<cache_t>& karg_)
        : karg(karg_)
    {
    }

    // template<int32_t acc_length>
    // __device__ static void precompute(float * acc){
    //     // compute per row square sum
    // }
    template <int32_t thread_length, int32_t segment_split>
    __device__ void update(cache_t* acc, emb_t* weight, uint32_t row_index)
    {
        if constexpr(segment_split == 0)
        {
            cache_t momentum = karg.p_momentum[row_index]; // should be s_load
            // compute per row square sum
            cache_t local_sum_squre = .0f;
            if constexpr(weigiht_decay_mode == 1)
            {
#pragma unroll
                for(auto i = 0; i < thread_length; i++)
                {
                    cache_t w = static_cast<cache_t>(weight[i]);
                    cache_t a = acc[i] + w * karg.weight_decay;
                    local_sum_squre += a * a;
                }
            }
            else
            {
#pragma unroll
                for(auto i = 0; i < thread_length; i++)
                {
                    cache_t a = acc[i];
                    local_sum_squre += a * a;
                }
            }

            cache_t avg_square =
                wave_reduce<reduce_op_sum_t<cache_t>, cache_t, AMDGCN_WAVE_SIZE>(local_sum_squre) /
                embedding_dim;

            cache_t multiplier;
            cache_t correction;

            cache_t momentum_new = momentum + avg_square;

            multiplier = karg.learning_rate / (sqrtf(momentum_new) + karg.eps);

            if constexpr(weigiht_decay_mode == 1)
            {
                correction = 1.0 - multiplier * karg.weight_decay;
            }
            else if constexpr(weigiht_decay_mode == 2)
            {
                correction = 1.0 - karg.learning_rate * karg.weight_decay;
            }
            else
            {
                correction = 1.0;
            }

// update new weight value
#pragma unroll
            for(auto i = 0; i < thread_length; i++)
            {
                cache_t w = static_cast<cache_t>(weight[i]);
                cache_t a = acc[i];
                w         = correction * w - multiplier * a;
                weight[i] = static_cast<emb_t>(w);
            }

            karg.p_momentum[row_index] = momentum_new;
        }
    }

    rowwise_adagrad_kernel_arg_t<cache_t> karg;
};

template <typename optimizer_t,
          typename optimizer_karg_t,
          typename emb_t,
          typename cache_t,
          typename grad_t,
          int32_t block_size,
          int32_t embedding_dim,
          int32_t bag_prefetch,
          int32_t bag_unroll,
          int32_t segment_split> // 0-warp per row, 1-cta per row, 2-atomic(needed?)
__device__ void split_tbe_backward_unweighted_hip_kernel(
    const grad_t* p_output_grad,
    emb_t* p_emb_table,
    const int64_t* p_sorted_linear_indices_run,
    const int64_t* p_sorted_linear_indices_cumulative_run_lengths,
    const int32_t* p_sorted_linear_indices_num_runs,
    const int32_t* p_long_run_ids,
    const int64_t* p_num_long_run_ids,
    const int32_t* p_sorted_infos,
    magic_div_u32_t batch_mdiv,
    uint32_t max_segment_length_per_warp,
    uint32_t emb_dim,
    uint32_t batch,
    uint32_t num_rows,
    uint32_t num_tables,
    optimizer_karg_t opt_karg)
{
    constexpr uint32_t dword_per_row   = (embedding_dim + THREADS_PER_ROW - 1) / THREADS_PER_ROW;
    constexpr uint32_t waves_per_block = block_size / AMDGCN_WAVE_SIZE;
    constexpr uint32_t length_mask     = ~(bag_unroll - 1);
    const uint32_t wave_id = __builtin_amdgcn_readfirstlane(threadIdx.x / AMDGCN_WAVE_SIZE);
    const uint32_t lane_id = threadIdx.x % AMDGCN_WAVE_SIZE;
    const uint32_t run_id  = wave_id + blockIdx.x * waves_per_block;

    if(run_id >= p_sorted_linear_indices_num_runs[0])
        return;

    const int64_t linear_index  = p_sorted_linear_indices_run[run_id];
    const int64_t emb_idx       = linear_index - blockIdx.y;
    const int32_t segment_start = p_sorted_linear_indices_cumulative_run_lengths[run_id];
    const int32_t segment_end   = p_sorted_linear_indices_cumulative_run_lengths[run_id + 1];

    p_output_grad += blockIdx.y * emb_dim;

    uint64_t emb_table_stride = static_cast<uint64_t>(num_rows) * emb_dim;
    p_emb_table += blockIdx.y * emb_table_stride;
    opt_karg.p_momentum += blockIdx.y * num_rows;

    const int32_t segment_length = segment_end - segment_start;

    if(segment_length >= max_segment_length_per_warp)
        return;

    const int32_t segment_length_mod = segment_length & length_mask;

    cache_t grad_acc[dword_per_row];
    int32_t infos[bag_unroll];
    grad_t grad_data[dword_per_row * bag_prefetch];
    emb_t emb_data[dword_per_row];

    int itr = 0;
    if(segment_length_mod == 0)
        goto L_tail_grad_acc;

#pragma unroll
    for(int i = 0; i < bag_unroll; i++)
    {
        infos[i] = p_sorted_infos[i];
    }

    itr += bag_unroll;
    p_sorted_infos += bag_unroll;

    uint32_t row_index;
    uint32_t table_index__unused;

    // LOOP
    for(; itr < segment_length_mod; itr += bag_unroll)
    {
        magic_div_u32_run_with_mod(batch_mdiv, infos[0], batch, table_index__unused, row_index);
        load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
            &grad_data[0], row_index * num_tables, p_output_grad, lane_id);

        magic_div_u32_run_with_mod(batch_mdiv, infos[1], batch, table_index__unused, row_index);
        load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
            &grad_data[dword_per_row], row_index * num_tables, p_output_grad, lane_id);

#pragma unroll
        for(int j = 2; j < bag_unroll; j += 2)
        {
            accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
                &grad_acc[0], &grad_data[0], lane_id);
            magic_div_u32_run_with_mod(batch_mdiv, infos[j], batch, table_index__unused, row_index);
            load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
                &grad_data[0], row_index * num_tables, p_output_grad, lane_id);

            accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
                &grad_acc[0], &grad_data[dword_per_row], lane_id);
            magic_div_u32_run_with_mod(
                batch_mdiv, infos[j + 1], batch, table_index__unused, row_index);
            load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
                &grad_data[dword_per_row], row_index * num_tables, p_output_grad, lane_id);
        }

#pragma unroll
        for(int i = 0; i < bag_unroll; i++)
        {
            infos[i] = p_sorted_infos[i];
        }
        p_sorted_infos += bag_unroll;

        accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
            &grad_acc[0], &grad_data[0], lane_id);
        accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
            &grad_acc[0], &grad_data[dword_per_row], lane_id);
    }

    // LAST
    magic_div_u32_run_with_mod(batch_mdiv, infos[0], batch, table_index__unused, row_index);
    load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
        &grad_data[0], row_index * num_tables, p_output_grad, lane_id);

    magic_div_u32_run_with_mod(batch_mdiv, infos[1], batch, table_index__unused, row_index);
    load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
        &grad_data[dword_per_row], row_index * num_tables, p_output_grad, lane_id);

#pragma unroll
    for(int j = 2; j < bag_unroll; j += 2)
    {
        accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
            &grad_acc[0], &grad_data[0], lane_id);
        magic_div_u32_run_with_mod(batch_mdiv, infos[j], batch, table_index__unused, row_index);
        load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
            &grad_data[0], row_index * num_tables, p_output_grad, lane_id);

        accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
            &grad_acc[0], &grad_data[dword_per_row], lane_id);
        magic_div_u32_run_with_mod(batch_mdiv, infos[j + 1], batch, table_index__unused, row_index);
        load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
            &grad_data[dword_per_row], row_index * num_tables, p_output_grad, lane_id);
    }

    accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
        &grad_acc[0], &grad_data[0], lane_id);
    accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
        &grad_acc[0], &grad_data[dword_per_row], lane_id);

L_tail_grad_acc:
    if(segment_length & (bag_unroll - 1))
    {
        // last, load one by one
        do
        {
            infos[0] = p_sorted_infos[0];
            p_sorted_infos++;

            magic_div_u32_run_with_mod(batch_mdiv, infos[0], batch, table_index__unused, row_index);
            load_row_per_warp<grad_t, embedding_dim, int32_t>::run(
                &grad_data[0], row_index * num_tables, p_output_grad, lane_id);
            accumulate_row_per_warp<grad_t, embedding_dim, cache_t>::run(
                &grad_data[0], &grad_data[0], lane_id);

            itr++;
        } while(itr < segment_length);
    }

    // load the old emb weight data
    load_row_per_warp<emb_t, embedding_dim, int32_t>::run(
        &emb_data[0], emb_idx, p_emb_table, lane_id);
    optimizer_t optimizer(opt_karg);
    optimizer.template update<dword_per_row, segment_split>(grad_acc, emb_data, row_index);

    // store updated weight to grad
    store_row_per_warp<emb_t, embedding_dim, emb_t>::run(&emb_data[0], p_emb_table, lane_id);
}

#define SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(optimizer,                                                                                 \
                                          weight_decay_mode,                                                                         \
                                          segment_split,                                                                             \
                                          emb_prec,                                                                                  \
                                          emb_type,                                                                                  \
                                          embedding_dim,                                                                             \
                                          bag_prefetch,                                                                              \
                                          bag_unroll)                                                                                \
    extern "C" __global__ void                                                                                                       \
        split_tbe_bwd_hip_kernel_warp_per_row_##optimizer##_w##weight_decay_mode##_s##segment_split##_##emb_prec##_e##embedding_dim( \
            const float* p_output_grad,                                                                                              \
            emb_type* p_emb_table,                                                                                                   \
            const int64_t* p_sorted_linear_indices_run,                                                                              \
            const int64_t* p_sorted_linear_indices_cumulative_run_lengths,                                                           \
            const int32_t* p_sorted_linear_indices_num_runs,                                                                         \
            const int32_t* p_long_run_ids,                                                                                           \
            const int64_t* p_num_long_run_ids,                                                                                       \
            const int32_t* p_sorted_infos,                                                                                           \
            magic_div_u32_t batch_mdiv,                                                                                              \
            uint32_t max_segment_length_per_warp,                                                                                    \
            uint32_t emb_dim,                                                                                                        \
            uint32_t batch,                                                                                                          \
            uint32_t num_rows,                                                                                                       \
            uint32_t num_tables,                                                                                                     \
            optimizer##_kernel_arg_t<float> opt_karg)                                                                                \
    {                                                                                                                                \
        split_tbe_backward_unweighted_hip_kernel<                                                                                    \
            optimizer##_optimizer_t<float, emb_type, embedding_dim, weight_decay_mode>,                                              \
            optimizer##_kernel_arg_t<float>,                                                                                         \
            emb_type,                                                                                                                \
            float,                                                                                                                   \
            float,                                                                                                                   \
            BLOCK_SIZE,                                                                                                              \
            embedding_dim,                                                                                                           \
            bag_prefetch,                                                                                                            \
            bag_unroll,                                                                                                              \
            segment_split>(p_output_grad,                                                                                            \
                           p_emb_table,                                                                                              \
                           p_sorted_linear_indices_run,                                                                              \
                           p_sorted_linear_indices_cumulative_run_lengths,                                                           \
                           p_sorted_linear_indices_num_runs,                                                                         \
                           p_long_run_ids,                                                                                           \
                           p_num_long_run_ids,                                                                                       \
                           p_sorted_infos,                                                                                           \
                           batch_mdiv,                                                                                               \
                           max_segment_length_per_warp,                                                                              \
                           emb_dim,                                                                                                  \
                           batch,                                                                                                    \
                           num_rows,                                                                                                 \
                           num_tables,                                                                                               \
                           opt_karg);                                                                                                \
    }

SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp32, float, 64, 2, 8)
SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp32, float, 128, 2, 8)
SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp32, float, 192, 2, 8)
SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp32, float, 256, 2, 8)

SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp16, half, 64, 2, 8)
SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp16, half, 128, 2, 8)
SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp16, half, 192, 2, 8)
SPLIT_TBE_BWD_WARP_PER_ROW_KERNEL(rowwise_adagrad, 1, 0, fp16, half, 256, 2, 8)