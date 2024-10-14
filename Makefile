.PHONY: winx64
winx64:
	GOOS=windows GOARCH=amd64 go build -o excelexporter.exe
