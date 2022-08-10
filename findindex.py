#coding=gbk

#匹配两个表的索引, 得到两个表的行数和图片的行数, 图片数量

import re
from openpyxl_image_loader import SheetImageLoader

pic_cols = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", )


def find(worksheet, picture_worksheet, follow_names, follow_cols):
	ret_dict = {"follow_line":[], "follow_index":{}, }
	max_row = worksheet.max_row
	for i in range(1, max_row + 1):
		is_find = False
		for col in follow_cols:
			cell = worksheet.cell(i, col)
			if cell == None:
				continue
			if cell.value in follow_names:
				is_find = True
				break
		if False == is_find:
			continue
		cell = worksheet.cell(i, 1)
		if cell == None:
			continue
		index = int(cell.value)
		ret_dict["follow_line"].append(i)
		follow_index_dict = ret_dict["follow_index"]
		follow_index_dict[index] = []

	img_loader = SheetImageLoader(picture_worksheet)
	pic_max_row = picture_worksheet.max_row
	follow_index_dict = ret_dict["follow_index"]
	for i in range(1, pic_max_row + 1):
		cell = picture_worksheet.cell(i, 1)
		if cell.value == None:
			continue
		match_obj = re.search("^\d+", str(cell.value))
		if match_obj == None:
			continue
		match_str = match_obj.group(0)
		index = int(match_str)

		if False == follow_index_dict.__contains__(index):
			continue

		next_row = 0
		for j in range(0, 50):
			now_row = i + j
			first_col_cell = picture_worksheet.cell(now_row, 1)
			if first_col_cell.value != None:
				if 0 == next_row:
					next_row = now_row
				else:
					break
			for col in pic_cols:
				coord = f'{col}{now_row}'
				if False == img_loader.image_in(coord):
					continue
				follow_index_dict[index].append(coord)

	return ret_dict

