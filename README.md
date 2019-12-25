# Barbara's Tool

Barbara wants to sync the vendor data to her spreadsheet. Most importantly, she has a programmer friend who happened to be me.

This is just a naive implementation base on requirements. No magic here!

I wish you happy every day, Barbara!

## Cross Compile to Windows on Mac

```shell
# Install MinGW which is required by CGO
brew install mingw-w64

# Cross compile
CGO_ENABLED=1 CC=x86_64-w64-mingw32-gcc CXX=x86_64-w64-mingw32-g++ GOOS=windows GOARCH=amd64 go build -ldflags="-H windowsgui"
```
