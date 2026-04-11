MATRIX="[]"
RUN_ANY="false"
IN_WIN="true"
IN_MAC="false"
IN_ARM="true"

if [ "$IN_WIN" == "true" ]; then 
  MATRIX=$(echo $MATRIX | jq '. += [{"os": "windows-latest", "suffix": "Windows"}]')
  RUN_ANY="true"
fi
if [ "$IN_MAC" == "true" ]; then 
  MATRIX=$(echo $MATRIX | jq '. += [{"os": "macos-15-intel", "suffix": "macOS-x86_64"}]')
  RUN_ANY="true"
fi
if [ "$IN_ARM" == "true" ]; then 
  MATRIX=$(echo $MATRIX | jq '. += [{"os": "macos-15", "suffix": "macOS-arm64"}]')
  RUN_ANY="true"
fi
echo "$MATRIX"
echo "run_any=$RUN_ANY"
