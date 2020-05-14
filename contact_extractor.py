"""!
This script is used to extract the contact information within an HTML file (exported from Telus webmail contact group)
and generate an Excel workbook with the found information.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, Message
import pandas as pd
import regex as re
from csv import DictWriter
import xlsxwriter
import os

name_regex = re.compile(r"(.*?)\s([A-z]{1}[A-z']+)$")
garbage_detection_regex = re.compile(r"[a-f0-9]{4}-[a-f0-9]{4}")

class userInterface:

    root: tk.Tk
    file_name: str = None
    tables: list = None
    contact_list: list = list()
    skewed_contact_list: list = list()

    def __init__(self, master):
        self.root = master
        self.root.title("HTML to XLSX")
        self.main_canvas = tk.Canvas(self.root, width = 300, height = 230, bg = 'deep sky blue', relief = 'raised')
        self.main_canvas.pack()
        self.add_elements()
        
    def add_elements(self):
        """!
        Adds all of the GUI elements to the main canvas.
        """

        self.title_label = tk.Label(self.root, text='Conversion Tool', bg = 'deep sky blue')
        self.title_label.config(font=('helvetica', 20))
        self.main_canvas.create_window(150, 30, window=self.title_label)
        self.button_chooseHTML = HoverButton(self.root, text="Choose HTML", command=self.open_HTML, bg='green2', activebackground='green3', fg='white', font=('helvetica', 12, 'bold'))
        self.main_canvas.create_window(150, 90, window=self.button_chooseHTML)  
        self.button_help = HoverButton(self.root, text='\u2753',command=self.open_help_window, bg='blue', activebackground='blue3', fg='white', font=('helvetica', 12, 'bold'))
        self.main_canvas.create_window(150, 130, window=self.button_help)
        self.button_exit = HoverButton(self.root, text='Exit Application',command=lambda: self.root.destroy(), bg='red', activebackground='red3',  fg='white', font=('helvetica', 12, 'bold'))
        self.main_canvas.create_window(150, 170, window=self.button_exit)

    def open_HTML(self):
        """!
        This is effectively the main function. It is tied to the 'Choose HTML' button.
        In this function, the HTML file is read into memory and then passed to other functions
        to parse the information.
        """

        try:
            self.cleanup()
            self.file_name = filedialog.askopenfilename(filetypes = (("HTML files","*.html"),))
            if ".html" in self.file_name:
                self.tables = pd.read_html(self.file_name)
                self.file_name = self.file_name.split(".html")[0]
            else:
                self.tables = pd.read_html(self.file_name+".html")
        except ImportError:
            messagebox.showerror("Error", "Unable to read file")
            return 1
        self.parse_HTML()
        flag = self.export_to_CSV() 
        if flag:
            messagebox.showerror("Error", "Conversion to CSV failed")
        else:
            self.convert_to_XLSX()
            messagebox.showinfo("Information","Contact list has been saved to an excel workbook at the following location:\n\n"+self.file_name+".xlsx"+"\n\nNumber of contacts found: "+str(len(self.contact_list)))
        print("Process finished.")

    def parse_HTML(self):
        """!
        Parse the HTML table to get the contact information.
        Generates the contact_list.
        """

        print("Analysing HTML page to find all contact entries.")

        try:
            contact_table = self.tables[3]
            print("Initializing empty contact list")
            contact_info = dict()

            for row in contact_table.iterrows():
                data = row[1][2]
                str_data = str(data)
                if str_data != "nan":
                    if str_data not in self.skewed_contact_list:
                        self.skewed_contact_list.append(str_data)
            if self.skewed_contact_list:
                messagebox.showwarning("Warning", "Contact information for the following people was not exported properly by Telus:\n\n"+"\n".join(self.skewed_contact_list)+"\n\nThe program will continue to run, but make sure you manually add the entries into the final excel workbook for these people.")

            for row in contact_table.iterrows():
                data = row[1][1]
                str_data = str(data)
                if str_data != "nan":
                    if str_data != "\"zoom yoga 2020\"":
                        if str_data[0].isupper():
                            if str_data == "SSUC":
                                if "Notes" not in contact_info.keys():
                                    garbage_match = re.search(garbage_detection_regex, str_data)
                                    if garbage_match is None:
                                        contact_info["Notes"] = str_data
                                else:
                                    garbage_match = re.search(garbage_detection_regex, str_data)
                                    if garbage_match is None:
                                        contact_info["Notes"] = contact_info["Notes"] + ", " + str_data
                                continue
                            if contact_info:
                                self.contact_list.append(contact_info)
                                print("Adding contact:", contact_info)
                                contact_info = dict()
                                match = re.search(name_regex, str_data)
                                if match is not None:
                                    first_name = match.group(1)
                                    last_name = match.group(2)
                                    print(first_name, last_name)
                                    contact_info["Name"] = last_name+", "+first_name
                                    contact_info["Last Name"] = last_name
                                else:
                                    raise("Failed to find a name in: "+str_data)
                            else:
                                match = re.search(name_regex, str_data)
                                if match is not None:
                                    first_name = match.group(1)
                                    last_name = match.group(2)
                                    print(first_name, last_name)
                                    contact_info["Name"] = last_name+", "+first_name
                                    contact_info["Last Name"] = last_name
                                else:
                                    raise("Failed to find a name in: "+str_data)
                        elif "@" in str_data:
                            contact_info["email"] = str_data
                        elif str_data.count('-') >= 2 and ":" not in str_data:
                            contact_info["phone"] = str_data
                        else:
                            if "Notes" not in contact_info.keys():
                                garbage_match = re.search(garbage_detection_regex, str_data)
                                if garbage_match is None:
                                    contact_info["Notes"] = str_data
                            else:
                                garbage_match = re.search(garbage_detection_regex, str_data)
                                if garbage_match is None:
                                    contact_info["Notes"] = contact_info["Notes"] + ", " + str_data

            self.contact_list.append(contact_info)
            print("Adding contact:", contact_info)
            print("Finished analysing the HTML. Found", len(self.contact_list), "contacts. Please confirm this with your online list")

            # Ensure there are no Null entries for "Notes"
            for contact in self.contact_list:
                if "Notes" not in contact:
                    contact["Notes"] = ""
        except Exception as e:
            self.show_exception(e)

    def export_to_CSV(self):
        """!
        Exports the generated contact list to a CSV file.
        """

        try:
            # print("Attempting to save to", self.file_name+".csv")
            with open(self.file_name+".csv", 'w', encoding='utf8', newline='') as output_file:
                fc = DictWriter(output_file, fieldnames=self.contact_list[0].keys(),)
                fc.writeheader()
                fc.writerows(self.contact_list)
        except IOError:
            messagebox.showerror("Error", "Unable to access "+ self.file_name + ".csv. Please close the file if it is open.")
            print("Unable to access", self.file_name+".csv.", "Please close the file if it is open.")
            return 1
        print("Contact list has been saved.")
        return 0

    def convert_to_XLSX(self):
        """!
        Converts the generated CSV file into an Excel workbook with the following customizations:
        1. First row is frozen
        2. Sorted by last name
        3. Removed column for last name and phone number
        """
        try:
            # Delete the xlsx file if it already exists
            if os.path.exists(self.file_name+".xlsx"):
                os.remove(self.file_name+".xlsx")

            read_file = pd.read_csv (self.file_name+".csv")
            read_file.sort_values('Last Name', inplace=True) # Sort by last name
            read_file.drop(columns=['Last Name'], inplace=True) # Remove last name column
            read_file.drop(columns=['phone'], inplace=True) # Remove phone number column
            writer = pd.ExcelWriter(self.file_name+".xlsx", engine='xlsxwriter')
            read_file.to_excel(writer, sheet_name="Contacts", index=False)  # send df to writer
            worksheet = writer.sheets["Contacts"]  # pull worksheet object
            for idx, col in enumerate(read_file):  # loop through all columns
                series = read_file[col]
                max_len = max((
                    series.astype(str).map(len).max(),  # len of largest item
                    len(str(series.name))  # len of column name/header
                    )) + 1  # adding a little extra space
                worksheet.set_column(idx, idx, max_len)  # set column width
            worksheet.freeze_panes(1, 0)  # Freeze the first row.
            writer.save()

            # Cleanup by deleting the csv that is not needed
            os.remove(self.file_name+".csv")
            
            return 0
        except Exception as e:
            self.show_exception(e)

    def cleanup(self):
        self.contact_list = list()
        self.file_name = ""
        self.tables = None
        self.skewed_contact_list = list()

    def open_help_window(self):
        """!
        Opens a new window at the top level of the GUI
        to display help information.
        """

        self.help_window = tk.Toplevel(self.root)
        message_str = ""
        message_str += "HOW TO:\n\n1. Open zoom address book\n"
        message_str += "2. Right click on zoom yoga group and select print. This will cause a new tab to open with all the contacts listed\n"
        message_str += "3. Close the print dialog but not the print page. Right click anywhere on the page, choose 'save as' and their should be an option for 'HTML complete webpage'\n"
        message_str += "4. Make sure the folder it is saving to is correct, if not, change to that folder\n"
        message_str += "5. Change the name to whatever you want (e.g, zoom yoga 2020 05 06) and save\n"
        message_str += "6. In this application, hit the 'choose HTML' button and select the HTML that was just saved and hit 'Open'\n"
        message_str += "7. The program will take the contacts from the HTML and create an excel workbook for you. Wait and observe any messages that pop-up.\n"
        message_str += "8. When the program finishes, close the information pop-up with 'ok' button\n"
        message_str += "9. Verify that the contacts within the workbook are correct, and add any that were flagged as incomplete. If you did not see any message pop-up during the run about improperly exported contacts, you can skip this.\n"
        w = Message(self.help_window, text=message_str)
        w.pack()

    def show_exception(self, e):
        messagebox.showerror("Exception", e)
        return 1

class HoverButton(tk.Button):
    """!
    Special button class with the following effects:
    1. Buttons change color on mouse-over events.
    """

    def __init__(self, master, **kw):
        tk.Button.__init__(self,master=master,**kw)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self['background'] = self['activebackground']

    def on_leave(self, e):
        self['background'] = self.defaultBackground

if __name__ == "__main__":
    master = tk.Tk()
    userInterface(master)
    master.mainloop()