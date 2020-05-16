build-raspberry:
	env GOOS=linux GOARCH=arm GOARM=7 go build -o xlsx-rasp main.go