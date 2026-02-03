package main

import (
	"fmt"
	"net/http/httptest"
	"os"

	"g3pix.com.br/axonasp/asp"
)

func main() {
	content, err := os.ReadFile("www/test_classes.asp")
	if err != nil {
		panic(err)
	}

	// Parse
	tokens := asp.ParseRaw(string(content))
	engine := asp.Prepare(tokens)

	// Run
	req := httptest.NewRequest("GET", "/", nil)
	w := httptest.NewRecorder()
	ctx := asp.NewExecutionContext(w, req, "www")

	engine.Run(ctx)

	fmt.Println(w.Body.String())
}
