@ECHO OFF
set GOARCH=amd64
set GOOS=windows
go build -ldflags -H=windowsgui
signtool.exe sign /fd SHA256 ddms-ar-converter.exe