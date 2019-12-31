# Barbara's Tool

Barbara wants to sync the vendor data to her spreadsheet. Most importantly, she has a programmer friend who happened to be me.

This is just a naive implementation base on requirements. No magic here!

I wish you happy every day, Barbara!

## Directories

```text
/
├──cmd                // Main application
│  └──gui             // GUI
└──pkg                // Libraries
   └──excel           // Excel Utilities
```

## Build

### Mac

```shell
# Build executable
# -s - Omit the symbol table and debug information.
# -w - Omit the DWARF symbol table.
go build -ldflags="-s -w"

# Compress executable
upx --ultra-brute BarbarasTool

# Remove terminal prompt
mv BarbarasTool BarbarasTool.app
```

### Cross Compile to Windows on Mac

```shell
# Install cross compiler MinGW
brew install mingw-w64

# Build executable
# -s            - Omit the symbol table and debug information.
# -w            - Omit the DWARF symbol table.
# -H windowsgui - On Windows, -H windowsgui writes a "GUI binary" instead of a "console binary."
CGO_ENABLED=1 CC=x86_64-w64-mingw32-gcc CXX=x86_64-w64-mingw32-g++ GOOS=windows GOARCH=amd64 go build -ldflags="-s -w -H windowsgui"

# Compress executable
upx --ultra-brute BarbarasTool.exe
```
