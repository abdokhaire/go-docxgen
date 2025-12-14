module github.com/abdokhaire/go-docxgen

go 1.23.2

require (
	github.com/Masterminds/sprig/v3 v3.3.0
	github.com/bep/imagemeta v0.8.1
	github.com/dlclark/regexp2 v1.11.4
	github.com/fumiama/go-docx v0.0.0-20250506085032-0c30fd09304b
	github.com/fumiama/imgsz v0.0.2
	github.com/stretchr/testify v1.10.0
	golang.org/x/image v0.21.0
	golang.org/x/text v0.19.0
)

replace github.com/fumiama/go-docx => ./pkg/go-docx

require (
	dario.cat/mergo v1.0.1 // indirect
	github.com/Masterminds/goutils v1.1.1 // indirect
	github.com/Masterminds/semver/v3 v3.3.0 // indirect
	github.com/davecgh/go-spew v1.1.1 // indirect
	github.com/google/uuid v1.6.0 // indirect
	github.com/huandu/xstrings v1.5.0 // indirect
	github.com/mitchellh/copystructure v1.2.0 // indirect
	github.com/mitchellh/reflectwalk v1.0.2 // indirect
	github.com/pmezard/go-difflib v1.0.0 // indirect
	github.com/shopspring/decimal v1.4.0 // indirect
	github.com/spf13/cast v1.7.0 // indirect
	golang.org/x/crypto v0.26.0 // indirect
	gopkg.in/yaml.v3 v3.0.1 // indirect
)
