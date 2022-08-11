#coding=gbk

import readbat

class InterfaceMode():
	character = 0
	picture = 1

config = readbat.run("config.bat")

def to_int(list):
	for i in range(len(list)):
		list[i] = int(list[i])
	return list

second_row = 2
third_row = 3
display_cols = to_int(config["display_cols"].split(","))
follow_cols = to_int(config["follow_cols"].split(","))
follow_names = config["follow_names"].split(",")

workbook = config["workbook"]
title = config["title"]
first_row = int(config["first_row"])
filter_col = int(config["filter_col"])
one_page_size = int(config["one_page_size"])
front_size = int(config["front_size"])
geometry = config["geometry"]
wraplength = config["wraplength"]
red_color = "#ff0000"
black_color = "#000000"



class Interface():
	def __init__(self):
		self.old_index = 0
		self.index = 0
		self.size = 0
		self.old_mode = 0
		self.mode = 0
		self.is_filter = False
		self.title_labels = []
		self.temporary_labels = []
		self.labelss = []
		self.filter_labelss = []
		self.old_labelss = self.labelss
		self.grid_labels_index = []
		self.last_time = 0

	def DisplayCharacter(self):
		old_labelss = self.old_labelss
		labelss = self.labelss if True != self.is_filter else self.filter_labelss
		self.size = len(old_labelss)
		old_first_index = self.old_index - self.old_index % one_page_size
		if self.old_index != self.index or self.old_mode != self.mode or old_labelss is not labelss:
			if self.old_index < self.size:
				old_labels = old_labelss[self.old_index]
				for label in old_labels:
					label.config(fg = black_color)

		if self.old_mode == InterfaceMode.picture:
			self.old_mode = InterfaceMode.character

			for label in self.temporary_labels:
				label.grid_remove()

		for index in self.grid_labels_index:
			if index < 0 or index >= self.size:
				continue
			labels = old_labelss[index]
			for label in labels:
				label.grid_remove()
		self.grid_labels_index.clear()

		if self.old_labelss is not labelss and self.index < self.size:
			new_labels = self.old_labelss[self.index]
			for label in new_labels:
				label.config(fg = black_color)
		self.size = len(labelss)
		if self.old_index >= self.size:
			self.old_index = 0
		if self.index >= self.size:
			self.index = 0

		if 0 == self.size:
			return
		new_labels = labelss[self.index]
		for label in new_labels:
			label.config(fg = red_color)

		first_index = self.index - self.index % one_page_size

		end_index = min(self.size, first_index + one_page_size)
		cols_size = len(display_cols)
		for i in range(first_index, end_index):
			if i < 0 or i >= self.size:
				continue
			labels = labelss[i]
			for j in range(len(labels)):
				if j >= cols_size:
					break
				label = labels[j]
				label.grid()
			self.grid_labels_index.append(i)
		self.old_labelss = labelss

	def DisplayPicture(self):
		if self.old_mode == InterfaceMode.character:
			self.old_mode = InterfaceMode.picture

		labelss = self.old_labelss
		if self.old_index < len(labelss):
			old_labels = labelss[self.old_index]
			for label in old_labels:
				label.config(fg = black_color)

		for index in self.grid_labels_index:
			if index < 0 or index >= self.size:
				continue
			labels = labelss[index]
			for label in labels:
				label.grid_remove()
		self.grid_labels_index.clear()

		if False == self.is_filter:
			labelss = self.labelss
		else:
			labelss = self.filter_labelss
		self.size = len(labelss)
		if self.old_index >= self.size:
			self.old_index = 0
		if self.index >= self.size:
			self.index = 0

		if 0 == self.size:
			return
		new_labels = labelss[self.index]
		cols_size = len(display_cols)
		has_pic = False
		for i in range(len(new_labels)):
			label = new_labels[i]
			if i < cols_size:
				temp_label = self.temporary_labels[i]
				temp_label["text"] = label.cget("text")
				temp_label.grid()
				continue
			label.grid()
			if False == has_pic:
				has_pic = True
		if True == has_pic:
			self.grid_labels_index.append(self.index)
		self.old_labelss = labelss

	def Display(self):
		if self.mode == InterfaceMode.character:
			self.DisplayCharacter()
		elif self.mode == InterfaceMode.picture:
			self.DisplayPicture()
	
	def left(self, event):
		self.old_index = self.index
		self.index -= one_page_size
		if self.index < 0:
			self.index = 0

	def right(self, event):
		self.old_index = self.index
		self.index += one_page_size
		if self.index >= self.size:
			if self.size > 0:
				self.index = self.size - 1
			else:
				self.index = 0

	def up(self, event):
		if self.index <= 0:
			return
		self.old_index = self.index
		self.index -= 1

	def down(self, event):
		if self.index >= self.size - 1:
			return
		self.old_index = self.index
		self.index += 1
	
	def picture_mode(self, event):
		if self.mode == InterfaceMode.character:
			self.mode = InterfaceMode.picture
	def character_mode(self, event):
		if self.mode == InterfaceMode.picture:
			self.mode = InterfaceMode.character

	def filter_mode(self, event):
		if self.is_filter == False:
			self.is_filter = True
			self.old_labelss = self.labelss
	def all_mode(self, event):
		if self.is_filter == True:
			self.is_filter = False
			self.old_labelss = self.filter_labelss

	def key(self, event, tk):
		if event.keysym == "Escape" or event.keysym == "q":
			tk.destroy()
			return

		elif event.keysym == "Up":
			self.up(event)
		elif event.keysym == "k":
			self.up(event)

		elif event.keysym == "Down":
			self.down(event)
		elif event.keysym == "j":
			self.down(event)

		elif event.keysym == "Left":
			self.left(event)
		elif event.keysym == "h":
			self.left(event)

		elif event.keysym == "Right":
			self.right(event)
		elif event.keysym == "l":
			self.right(event)

		elif event.keysym == "Return":
			self.picture_mode(event)
		elif event.keysym == "i":
			self.picture_mode(event)

		elif event.keysym == "BackSpace":
			self.character_mode(event)
		elif event.keysym == "o":
			self.character_mode(event)

		elif event.keysym == "f":
			self.filter_mode(event)
		elif event.keysym == "a":
			self.all_mode(event)

		else:
			return

		self.Display()
