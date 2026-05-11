package main

import (
	"fmt"
	"g3pix.com.br/axonasp/axonvm"
)

func main() {
	fmt.Printf("OpJSSetProto: %d\n", axonvm.OpJSSetProto)
	fmt.Printf("OpJSSuperCall: %d\n", axonvm.OpJSSuperCall)
	fmt.Printf("OpJSRot: %d\n", axonvm.OpJSRot)
}
