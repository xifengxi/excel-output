#!/bin/bash

CGO_ENABLED=0 GOOS=darwin GOARCH=amd64 go build -o diffexcel main.go
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o diffexcel.exe main.go