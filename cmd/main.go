// #cgo CFLAGS: -O3
// #cgo CXXFLAGS: -O3
package main

import (
	"github.com/andlabs/ui"
	"github.com/limingjie/BarbarasTool/cmd/gui"
)

func main() {
	ui.Main(gui.NewGUI)
}
