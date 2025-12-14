package functions

import "text/template"

// DefaultFuncMap is empty by default. Users can register their own functions
// using RegisterFunction or RegisterFuncMap (e.g., with Sprig).
var DefaultFuncMap = template.FuncMap{}
