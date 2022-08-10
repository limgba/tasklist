#coding=gbk

def run(file_name):
	file = open(file_name, "r", encoding="utf-8")
	ret = {}
	for str in file:
		index = str.find("set ")
		if -1 == index:
			continue
		str = str.replace("set ", "")
		str = str.replace(" ", "")
		str = str.replace("\"", "")
		str = str.replace("\n", "")

		index = str.find("=")
		key = str[:index]
		value = str[index + 1:]
		ret[key] = value
	file.close()
	return ret
