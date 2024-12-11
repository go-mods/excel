module github.com/go-mods/excel

go 1.22

toolchain go1.23.3

require (
	github.com/go-mods/convert v0.5.0
	github.com/go-mods/tags v1.1.3
	github.com/stretchr/testify v1.10.0
	github.com/xuri/excelize/v2 v2.8.0
)

require (
	github.com/davecgh/go-spew v1.1.1 // indirect
	github.com/dromara/carbon/v2 v2.5.2 // indirect
	github.com/mohae/deepcopy v0.0.0-20170929034955-c48cc78d4826 // indirect
	github.com/pmezard/go-difflib v1.0.0 // indirect
	github.com/richardlehane/mscfb v1.0.4 // indirect
	github.com/richardlehane/msoleps v1.0.4 // indirect
	github.com/xuri/efp v0.0.0-20241211021726-c4e992084aa6 // indirect
	github.com/xuri/nfp v0.0.0-20240318013403-ab9948c2c4a7 // indirect
	golang.org/x/crypto v0.30.0 // indirect
	golang.org/x/net v0.32.0 // indirect
	golang.org/x/text v0.21.0 // indirect
	gopkg.in/yaml.v3 v3.0.1 // indirect
)

replace github.com/xuri/excelize/v2 => github.com/go-mods/excelize/v2 v2.0.0-20231116122542-ce766d7021db
