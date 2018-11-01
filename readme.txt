1, excel2json excelFile 
转换当前目录下excel文件到json

2, excel2json outDir allSheet needDataType
转换当前目录下所有excel文件到目标目录，并且转换所有sheet和需要数据类型

3, excel2json srcDir outDir allSheet needDataType
转换源目录下所有excel文件到目标目录，并且转换所有sheet和需要数据类型

4, excel2json jsonFile
转换一个json文件到csv文件中


excel格式

第一行，标记用

A1：
	[uniqueid] 表示当前表使用唯一id
		格式：{sheet:{key:{line}}}
	[repeatid] 表示当前表使用可重复id
		格式：{sheet:{key:[{line1}, {line2}]}}
	
第一列数据格式只能是数字，其他列的格式可以在第一行标记出来
	空字符串或者s 表示本列是字符串类型
	f 表示本列是浮点类型
	i 表示本列是整形

第二行 是中文标记，仅仅用于标注
第三行 是程序中使用的关键字，必须是英文，使用驼峰风格，如:TemplateId

sheet名 使用驼峰命名，必须是英文
