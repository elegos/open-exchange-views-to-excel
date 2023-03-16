# Simple Open Exchange to Excel converter

Converts an Open Exchange file (exported via Archi) into a readable XLSX, containing the Elements of each view (name, documentation).

## Arguments

- `--input, -i`: the file to process. If not specified, all xml files inside input/
- `--output-dir, -o`: the folder where to output the excel files. If not specified, all xlsx files will be generated in output/
- `--view-filter, -f`: if specified, only the views which name match the passed filter (regex) will be output in the excel file

## Extras

- `archi-to-excel` is a script which can be installed in `~/.local/bin`. It will call the wrapper, letting the script to specify a local file and including the `-o` option automatically as the current directory, for example `archi-to-excel -i ./my_file.xml`. Note: change the path variable `script_dir`.