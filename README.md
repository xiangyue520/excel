# excel
将一个很大的excel拆分为一个个小的excel


## 准备
需要将excel 文件授予执行权限,不同操作系统选择不同目录下的执行文件,且为64位

## 使用说明
```markdown
./excel /tmp/a.xlsx 2 c Sheet1 标题行个数
第一个参数为excel文件地址,全路径
第二个参数为每页切分多少数据,默认为5000
第三个参数为切分后的文件名称前缀,在该例中生成的为/tmp/a_archive的文件夹,文件名称为`c_累加的个数`
第四个参数为需要查询的sheet名称,默认为Sheet1
第四个参数为保存的标题行个数,默认为1,即每个excel的都会包含的头部信息

eg1:如果想要拆分文件,5000一个表格,文件名和传入名称一致,且sheet名称为Sheet1
./excel /tmp/a.xlsx
eg2:如果想要拆分文件,2000一个表格,文件名和传入名称一致,且sheet名称为Sheet1
./excel /tmp/a.xlsx 2000

eg3:如果想要拆分文件,2000一个表格,文件名为d,且sheet名称为Sheet1
./excel /tmp/a.xlsx 2000 d

eg4:如果想要拆分文件,2000一个表格,文件名为d,且sheet名称为aa
./excel /tmp/a.xlsx 2000 d aa

eg4:如果想要拆分文件,2000一个表格,文件名为d,且sheet名称为Sheet1,每个excel保留2个标题行
./excel /tmp/a.xlsx 2000 d aa 2
```

## 编辑
```markdown
mac m1 aarch64编译
CGO_ENABLED=0 GOOS=darwin GOARCH=arm64  go build -o dist/macos/excel

windows 64位编译
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o dist/windows/excel.exe


linux x86_64编译
CGO_ENABLED=0 GOOS=linux GOARCH=amd64 go build -o dist/linux/excel


同时编译多个平台,需要安装gox
go get github.com/mitchellh/gox

gox -osarch="darwin/arm64 linux/amd64 windows/amd64" excel.go

```