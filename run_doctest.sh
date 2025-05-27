#!/bin/bash

# Run the doctest multiple times
for i in {1..10}; do
  echo "Run $i:"
  mix test "lib/umya_spreadsheet.ex:1363" --only doctest
  echo ""
done
