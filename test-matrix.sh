MATRIX="[]"
MATRIX=$(echo $MATRIX | jq '. += [{"os": "windows-latest", "suffix": "Windows"}]')
echo "matrix={\"include\":$MATRIX}"
