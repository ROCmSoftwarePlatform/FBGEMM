#!/bin/bash

results_file="test_results.log"
error_logs_dir="error_logs"
mkdir -p "$error_logs_dir"
echo "Test Suite,Test Case,Status,Time,Warnings,Errors,Skipped" > "$results_file"

run_python_test () {
  local test_file="$1"
  local test_name
  test_name=$(basename "$test_file" .py)
  local error_log_file="${error_logs_dir}/${test_name}_error.log"

  echo "################################################################################"
  echo "# Running Test Suite: ${test_file}"
  echo "################################################################################"

  output=$(python -m pytest -v -rsx -s -W ignore::pytest.PytestCollectionWarning \
    --timeout_method thread --timeout=600 --cache-clear "${test_file}" 2>&1)
  if echo "$output" | grep -q "FAILURES\|FAILED"; then
    suite_status="FAIL"
  elif echo "$output" | grep -q "ERROR"; then
    suite_status="ERROR"
  elif echo "$output" | grep -q "SKIPPED"; then
    suite_status="SKIPPED"
  else
    suite_status="PASS"
  fi
  if [ "$suite_status" = "FAIL" ] || [ "$suite_status" = "ERROR" ]; then
    echo "$output" > "$error_log_file"
  fi
  summary_line=$(echo "$output" | grep -E "^=+.*(error|errors|passed|failed|skipped|warning|warnings?).*in [0-9]+(\.[0-9]+)?s.*=+$" | tail -1)

  if [ -z "$summary_line" ] && [ -f "$error_log_file" ]; then
    summary_line=$(grep -E "^=+.*(error|errors|passed|failed|skipped|warning|warnings?).*in [0-9]+(\.[0-9]+)?s.*=+$" "$error_log_file" | tail -1)
  fi
  if [ -z "$summary_line" ]; then
    total_time="0s"
    passed="0"
    failed="0"
    warnings="0"
    errors="0"
    skipped="0"
  else
    total_time=$(echo "$summary_line" | grep -oP '\d+\.\d+s|\d+s' || echo "0s")
    passed=$(echo "$summary_line" | grep -oP '(\d+) passed' | grep -oP '\d+' || echo "0")
    failed=$(echo "$summary_line" | grep -oP '(\d+) failed' | grep -oP '\d+' || echo "0")
    warnings=$(echo "$summary_line" | grep -oP '(\d+) warnings?' | grep -oP '\d+' || echo "0")
    errors=$(echo "$summary_line" | grep -oP '(\d+) errors?' | grep -oP '\d+' || echo "0")
    skipped=$(echo "$summary_line" | grep -oP '(\d+) skipped' | grep -oP '\d+' || echo "0")
    if [ "$errors" -eq "0" ]; then
      errors=$(echo "$summary_line" | grep -oP '(\d+) error' | grep -oP '\d+' || echo "0")
    fi
    if [ "$warnings" -eq "0" ]; then
      warnings=$(echo "$summary_line" | grep -oP '(\d+) warning' | grep -oP '\d+' || echo "0")
    fi
  fi
  echo "$output" | grep -E "^(.*::.*::.*) (PASSED|FAILED|ERROR|SKIPPED)" | while read -r line; do
    test_case=$(echo "$line" | awk '{print $1}')
    test_status=$(echo "$line" | awk '{print $2}')
    test_time=$(echo "$line" | grep -oP "\(\d+\.\d+s\)" | tr -d '()' || echo "N/A")

    echo "$test_file,$test_case,$test_status,$test_time,$warnings,$errors,$skipped" >> "$results_file"
  done
  echo "$test_file,Suite Summary,$suite_status,$total_time,$warnings,$errors,$skipped" >> "$results_file"
  echo "$output"
}

echo "Starting all tests in directory: fbgemm_gpu/test"
for test_file in $(find . -name '*_test.py'); do
  run_python_test "$test_file"
done

python3 generate_excel_report.py
