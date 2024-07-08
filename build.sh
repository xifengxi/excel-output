#!/bin/bash

CGO_ENABLED=0 GOOS=darwin GOARCH=amd64 go build -o diffexecl main.go
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o diffexecl.exe main.go