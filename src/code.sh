#!/bin/bash

set -exo pipefail

main() {

    echo "Value of tsv_input: '$tsv_input'"
    echo "Value of output_filename: '$output_filename'"

    # Download the provided MetricsOutput.tsv file
    dx download "$tsv_input" -o tsv_input

    # Install all python packages from this app
    sudo -H python3 -m pip install --no-index --no-deps packages/*

    # If python script will run without output_filename string if it is provided
    if [ -z "$output_filename" ];
    then
        python3 process_metrics_file.py tsv_input
    else
        python3 process_metrics_file.py tsv_input -o "$output_filename"
    fi

    # Find output file from python script
    excel_file=$(find . -name "*.xlsx")

    echo File "$excel_file" created.

    # Upload excel file
    output_file=$(dx upload "$excel_file" --brief)

    dx-jobutil-add-output output_file "$output_file" --class=file
}
