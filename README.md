 # embed_file_in_vba_macro

Python script that takes in a file and generates an Office VBA macro that can regenerate the input file. Can be used to evade content analysis.

### Steps

1) Run script with your input file (e.g. `malware.exe`) as first argument:

```
python3 exe_to_vba_macro.py malware.exe
```

Example output:

```
ohsyln:~/projects/embed_vba$ python3 exe_to_vba_macro.py malware.exe
Input file read successfully: malware.exe (3164 bytes)
VBA macro output to: output.txt
```
2) Copy contents of output.txt into a macro in a Word doc

3) Macro runs and drops your input file (`malware.exe`) as `vba_gen.txt` whenever the Word doc is opened 

## License

MIT
