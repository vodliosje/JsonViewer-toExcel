import sys, json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
	QTableWidgetItem, QHeaderView, QPushButton, QLabel, QLineEdit, QFileDialog,
	QMessageBox, QTabWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QIcon 
from collections import defaultdict

class ViewJson(QWidget):
	def __init__ (self):
		super().__init__()
		self.setWindowTitle('View Json File')
		self.setGeometry(700,200,600,300)
		self.setWindowIcon(QIcon("app_pic.png"))

		self.ask_button = QPushButton("Browse",self)
		self.export_button = QPushButton("Export",self)
		self.refresh_button = QPushButton("Refresh",self)
		self.path_entry = QLineEdit("G:/laspongologist/froggie/excel101/basic_data.json",self)

		self.all_tabs = {}

		self.initUI()
	def initUI(self):

		self.vbox = QVBoxLayout(self)
		self.setLayout(self.vbox)

		heading =  QLabel('Choose the Json File',self)
		self.vbox.addWidget(heading)
		self.vbox.addWidget(self.path_entry)

		hbox = QHBoxLayout()
		hbox.addWidget(self.ask_button)
		hbox.addWidget(self.export_button)
		hbox.addWidget(self.refresh_button)
		self.vbox.addLayout(hbox)
		self.vbox.setAlignment(heading, Qt.AlignCenter)

		self.tabs = QTabWidget()

		self.ask_button.clicked.connect(self.ask_json_file)
		self.export_button.clicked.connect(self.ask_location)
		self.refresh_button.clicked.connect(self.refresh)

		self.setStyleSheet("""
			QLabel{
			font-size: 30px;
			font-family: Montserrat;
			font-weight: bold;
			border-radius:20px;
			}
			QMessageBox QLabel{
			font-size: 16px;
			font-weight: normal;
			font-family: Calibri;
			}

			""")

	def ask_json_file(self):
		
		self.json_path, _ = QFileDialog.getOpenFileName(
			self,
			"Select Json File", 
			f"{self.path_entry.text()}",
			"JSON Files (*.json);; All Files(*)")
		
		#print(f"{type(json_path)} of {json_path}")
		self.digest_json_file(self.json_path)

	def digest_json_file(self,path):
		#print(f'eating {path}')

		#-----Get and set the column-----------
		tabs = set()
		
		all_keys = {}
		row_num = ()

		global data
		with open (path, 'r') as t:
			data = json.load(t)
		
		for document, content in data.items():
			#print (document)
			tabs.update(content.keys())
			for key, category in content.items():
				if key not in all_keys:
					all_keys[key] = set()
				all_keys[key].update(category.keys())
			row_num = row_num + (document,)

		self.launch_data(row_num, tabs, all_keys)

	def launch_data(self,row_num, tabs, all_keys):
		
		#-------Set up the table-----------------
		self.vbox.addWidget(self.tabs)
		for tab_name in sorted(tabs):
			table = QTableWidget(len(row_num), len(all_keys[tab_name])+1)
			table.setHorizontalHeaderItem(0,QTableWidgetItem("name"))
			for i, header in enumerate(all_keys[tab_name]):
				table.setHorizontalHeaderItem(i+1,QTableWidgetItem(header))
			table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

			self.tabs.addTab(table, tab_name)
			self.all_tabs[tab_name] = table

		#-------------------Add data to the table-------------------
			for row_idx, person in enumerate(row_num):
				self.all_tabs[tab_name].setItem(row_idx,0,QTableWidgetItem(person))
				
				col_idx = 1
				for cell_data in all_keys[tab_name]:
					self.all_tabs[tab_name].setItem(row_idx,col_idx,QTableWidgetItem(data[person][tab_name][cell_data]))
					col_idx += 1 

	def ask_location (self):
		
		export_path, _ = QFileDialog.getSaveFileName(
			self,
			"Select JSON Library FIle",
			f"exported_data_{QDate.currentDate().toString(Qt.ISODate)}.xlsx",
			"Excel Files (*.xlsx);; CSV Files (*.csv)") 
		if not export_path.lower().endswith('.xlsx' or '.csv'):
			export_path += '.xlsx' 

		self.export_json(export_path)

	def export_json (self,export_path):
		
			digested_data = defaultdict(list)
			for name, tab in data.items():
				for tab_name, cell_data in tab.items():
					record = {}
					record['name'] = name 

					record.update(cell_data)
					digested_data[tab_name].append(record)

			self._export_to_excel(digested_data,export_path)
		
	def _export_to_excel(self,categorized_data, excel_file_path):
	    
	    # The engine='openpyxl' is the standard for modern .xlsx files
	    writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')
	    
	    # Check if any data was processed
	    if not categorized_data:
	        print("No valid data found for export.")
	        return

	    for category_name, records_list in categorized_data.items():
	        
	        df = pd.DataFrame(records_list)
	        df.to_excel(writer, sheet_name=category_name, index=False)
	        
	        #print(f"Wrote {len(records_list)} rows to sheet: '{category_name}'")

	    writer.close()
	    QMessageBox.information(self, "Save success", f"\nSuccessfully exported data to: {excel_file_path}")

	def refresh (self):
		self.vbox.removeWidget(self.tabs)
		self.tabs = None
		self.all_tabs = {}
		pass

def main():
	app = QApplication(sys.argv)
	window = ViewJson()
	window.show()
	sys.exit(app.exec_())

if __name__ == "__main__":
	main()