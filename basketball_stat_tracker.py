from tkinter import *
import xlsxwriter
import os

def output_stats(stats):
	workbook = xlsxwriter.Workbook('hello_world.xlsx')
	worksheet = workbook.add_worksheet()
	# set up basics
	quarters = ["1st Quarter", "2nd Quarter", "3rd Quarter", "4thQuarter", "Over Time"]
	stat_category = ["2 pt made", "2 pt attempt", "3 pt made", "3 pt attempt", 
	"free throw made", "free throw attempt","Off. rebound", "Def. rebound",
	"T.O.","Assist", "Block", "Steal","Your Score", "Opp Score"]
	worksheet.write_row('B1', quarters)
	worksheet.write_column('A2', stat_category)
	# load data
	for key in range(5):
		worksheet.write_column(1,key+1,stats[key])
	workbook.close()
	print("finish")
 

class App(Tk):
	def __init__(self, *args, **kwargs):
		Tk.__init__(self, *args, **kwargs)

		#Setup Frame
		container = Frame(self)
		container.pack(side="top", fill="both", expand=True)
		container.grid_rowconfigure(0, weight=1)
		container.grid_columnconfigure(0, weight=1)

		self.frames = {}
		for F in (FirstQuarter, SecondQuarter, ThirdQuarter, FourthQuarter, OverTime):
			frame = F(container, self) #call each frame
			self.frames[F] = frame
			frame.grid(row=0, column=0, sticky="nsew")

		self.show_frame(FirstQuarter)	
	def show_frame(self, context):
		frame = self.frames[context]
		frame.tkraise()
	def set_stats(self, new_frame):
		content = {"Shooting:": ["2 pt made", "2 pt attempt", "3 pt made", "3 pt attempt", 
		"free throw made", "free throw attempt"],
		"Rebounds:" : ["Off. rebound", "Def. rebound"],
		"Others:" : ["T.O.","Assist", "Block", "Steal"]}
		order = ["Shooting:", "Rebounds:","Others:"]
		attr_name = []
		for column_index, column in enumerate(order):
			column_name = column+"_name"
			column_name = Label(new_frame, text=column)
			column_name.grid(row = 1, column = 2*column_index, sticky = "E")
			for row_index, name in enumerate(content[column]):
				if name == "":
					continue
				#label_name = name+"_label"
				#spin_name = name+"_spin"
				label_name = Label(new_frame, text = name)
				label_name.grid(row=row_index+2,column=column_index*2,sticky="E",pady=5,padx=5)
				spin_name = Spinbox(new_frame, from_=0, to=100, width=5)
				spin_name.grid(row=row_index+2,column=column_index*2+1,sticky="E",pady=5,padx=5)
				attr_name.append(spin_name)

		your_score = Label(new_frame, text="Your Score")
		your_score.grid(row = 6, column = 4, sticky = "E")
		your_entry = Entry(new_frame, bd =1,width=7)
		your_entry.grid(row = 6, column = 5,sticky="E",pady=5,padx=5)
		attr_name.append(your_entry)
		opp_score = Label(new_frame, text="Opp. Score")
		opp_score.grid(row = 7, column = 4, sticky="E")
		opp_entry = Entry(new_frame, width=7)
		opp_entry.grid(row=7, column=5,sticky="E",pady=5,padx=5)
		attr_name.append(opp_entry)

		return attr_name
	
	def gather_stats(self):
		total_stats = {}
		quarters = (FirstQuarter, SecondQuarter, ThirdQuarter, FourthQuarter, OverTime)
		for index, F in enumerate(quarters):
			quarter_stats = []
			attr_list = self.frames[F].attr_list
			for attr in attr_list:
				if attr.get() == "":
					quarter_stats.append(0)
				else:
					quarter_stats.append(int(attr.get()))
				#quarter_stats.append(int(attr.get()))
			total_stats[index]=quarter_stats
		output_stats(total_stats)

		
class FirstQuarter(Frame):
	def __init__(self, parent, controller):
		#parent is the container, controller is self, which is the App
		Frame.__init__(self, parent)
		label = Label(self, text="First Quarter")
		label.grid(row=0, column=0)
		page_two = Button(self, text="2nd Quarter", command=lambda:controller.show_frame(SecondQuarter))
		page_two.grid(row=10, column=5,pady=5,padx=5)
		self.attr_list = controller.set_stats(self)


class SecondQuarter(Frame):
	def __init__(self, parent, controller):
		Frame.__init__(self, parent)

		label = Label(self, text="First Quarter")
		label.grid(row=0, column=0)
		page_one = Button(self, text="1st Quarter", command=lambda:controller.show_frame(FirstQuarter))
		page_one.grid(row=10, column=4,pady=5,padx=5)
		page_two = Button(self, text="3rd Quarter", command=lambda:controller.show_frame(ThirdQuarter))
		page_two.grid(row=10, column=5,pady=5,padx=5)
		self.attr_list =controller.set_stats(self)

class ThirdQuarter(Frame):
	def __init__(self, parent, controller):
		Frame.__init__(self, parent)

		label = Label(self, text="Third Quarter")
		label.grid(row=0, column=0)
		page_one = Button(self, text="2nd Quarter", command=lambda:controller.show_frame(SecondQuarter))
		page_one.grid(row=10, column=4,pady=5,padx=5)
		page_two = Button(self, text="4th Quarter", command=lambda:controller.show_frame(FourthQuarter))
		page_two.grid(row=10, column=5,pady=5,padx=5)
		self.attr_list =controller.set_stats(self)

class FourthQuarter(Frame):
	def __init__(self, parent, controller):
		Frame.__init__(self, parent)

		label = Label(self, text="Fourth Quarter")
		label.grid(row=0, column=0)
		page_one = Button(self, text="3rd Quarter", command=lambda:controller.show_frame(ThirdQuarter))
		page_one.grid(row=10, column=4,pady=5,padx=5)
		page_two = Button(self, text="Over Time", command=lambda:controller.show_frame(OverTime))
		page_two.grid(row=10, column=5,pady=5,padx=5)
		self.attr_list =controller.set_stats(self)

class OverTime(Frame):
	def __init__(self, parent, controller):
		Frame.__init__(self, parent)

		label = Label(self, text="Over Time")
		label.grid(row=0, column=0)
		page_one = Button(self, text="4th Quarter", command=lambda:controller.show_frame(FourthQuarter))
		page_one.grid(row=10, column=4,pady=5,padx=5)
		page_two = Button(self, text="Finish", command=lambda:controller.gather_stats())
		page_two.grid(row=10, column=5,pady=5,padx=5)
		self.attr_list =controller.set_stats(self)


app = App()
app.mainloop()