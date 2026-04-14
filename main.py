import pandas as pd
import io
import re
import csv
import processor
from datetime import datetime
import os
import tkinter.messagebox as msgbox
import threading
import tkinter as tk

class Controller:
    def __init__(
        self,
        data,
    ):
        self.data=data
    def show_error(self, message):
        """Central GUI error popup."""
        msgbox.showerror("Error", message)


    def validate_inputs(self):
        """Validate form inputs intelligently based on checkbox selections."""
        chosendate = self.data.get("start_date", "").strip()
        chosentime = self.data.get("start_hour", "").strip()
        chosend = self.data.get("end_date", "").strip()
        chosent = self.data.get("end_hour", "").strip()
        csv_path = self.data.get("csv_path", "").strip('"')
        dest_path = self.data.get("dest_path", "").strip('"')
        new_folder = self.data.get("new_folder", "").strip('"')
        use_dest = self.data.get("use_dest", False)
        use_create = self.data.get("use_create", False)

        invalid_chars = r'<>"|?*'

        # Always validate Date (YY/MM/DD)
        try:
            datetime.strptime(chosendate, "%d/%m/%y")
        except ValueError:
            self.show_error(
                f"Invalid date format: '{chosendate}'. Expected YY/MM/DD (e.g. 25/09/25)"
            )
            return False

        # ✅ Always validate Time (HH:MM)
        try:
            datetime.strptime(chosentime, "%H:%M")
        except ValueError:
            self.show_error(
                f"Invalid time format: '{chosentime}'. Expected HH:MM (e.g. 08:30)"
            )
            return False
        
        if chosend:
            try:
                datetime.strptime(chosend, "%d/%m/%y")
            except ValueError:
                self.show_error(
                    f"Invalid date format: '{chosend}'. Expected YY/MM/DD (e.g. 25/09/25)"
                )
                return False

        if chosent:
            try:
                datetime.strptime(chosent, "%H:%M")
            except ValueError:
                self.show_error(
                    f"Invalid time format: '{chosent}'. Expected HH:MM (e.g. 08:30)"
                )
                return False

        # CSV Path must always exist
        if not os.path.exists(csv_path):
            self.show_error(f"CSV file not found:\n{csv_path}")
            return False

        # Destination Path — only if 'Use Destination Path' is ticked
        if use_dest:
            if not dest_path:
                self.show_error("Please specify a destination path.")
                return False

            
            if not os.path.exists(dest_path):
                self.show_error(f"Destination folder does not exist:\n{dest_path}")
                return False

        # New Folder/Name — only if 'Create New Folder/Name' is ticked
        if use_create:
            if not new_folder:
                self.show_error("Please enter a new folder name.")
                return False
            if not new_folder.strip():
                self.show_error("Folder name cannot be empty or only spaces.")
                return False
            if not os.path.exists(new_folder):
                self.show_error(f"Destination folder does not exist:\n{new_folder}")
                return False
        
        for label, value in [("Threshold 1", self.data.get("threshold1", "").strip()),("Threshold 2", self.data.get("threshold2", "").strip())]:
            if value == "":
                continue
            try:
                float(value)
            except ValueError:
                self.show_error(f"{label} must be numeric (int or float). You entered: '{value}'")
                return False


        return True 
    

    def load_csv_file(self, file_path):
        """Attempt to open and read a CSV file safely."""
        rows = []
        max_len = 0
        ILLEGAL_CHARACTERS_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]")

        try:
            with open(file_path, encoding="utf-8", errors="ignore") as f:
                cleaned_lines = (ILLEGAL_CHARACTERS_RE.sub("", line) for line in f)
                reader = csv.reader(cleaned_lines)
                for row in reader:
                    rows.append(row)
                    max_len = max(max_len, len(row))

            if not rows or len(rows[0]) < 2:
                self.show_error("This file does not appear to be a valid CSV or contains no readable data.")
                return False

            return True

        except FileNotFoundError:
            self.show_error(f"File not found:\n{file_path}")
        except PermissionError:
            self.show_error(f"Permission denied when trying to read:\n{file_path}")
        except csv.Error as e:
            self.show_error(f"CSV parsing error:\n{e}")
        except Exception as e:
            self.show_error(f"Could not read file:\n{file_path}\n\nError details:\n{e}")

        return False





    

    def run(self):
        print("running")
        """Main entry point."""
        if not self.validate_inputs():
            return  #user already saw message
        
        file_path=self.data.get("csv_path", "").strip()
        if not self.load_csv_file(file_path):
            return

        #proceed wth CSV logic
        msgbox.showinfo("Processing", "Inputs valid! Processing started...")
        headers = ["Date/Time", "Test name", "Current user", "Position", "Serial", "Trigger point", "Results"]
        rows = []
        max_len = 0
        ILLEGAL_CHARACTERS_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]")

        with open(file_path, encoding="utf-8", errors="ignore") as f:
                cleaned_lines = (ILLEGAL_CHARACTERS_RE.sub("", line) for line in f)
                reader = csv.reader(cleaned_lines)
                for row in reader:
                    rows.append(row)
                    max_len = max(max_len, len(row))
        

        # pad uneven rows
        fixed_rows = [row + [""] * (max_len - len(row)) for row in rows]

        # insert custom header
        custom_header = ["Date", "YY/MM/DD", "Time", "HH/MM/SS", "holder1", "holder2"]
        fixed_rows.insert(0, custom_header)

        # write to buffer for pandas
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerows(fixed_rows)
        output.seek(0)
        df = pd.read_csv(output)
        
        print(df.head())


        column2 = df["YY/MM/DD"]
        column4 = df["HH/MM/SS"]

            
        chosendate = self.data["start_date"]
        chosentime = self.data["start_hour"]
        date_time = (chosendate, chosentime)




        chosend = self.data.get("end_date", "").strip()
        chosent = self.data.get("end_hour", "").strip()
        srow_index = processor.getstartrow(chosendate, chosentime, column2, column4)

        if srow_index == None:
            return
        if chosend and chosent:
            erow_index = processor.getstoprow(chosend, chosent, column2, column4)
        else:
            erow_index = None

        dictndmissingvalues = processor.createcsvdict(file_path, srow_index,erow_index)
        dictionary = dictndmissingvalues[0]

        xclfile = self.data.get("dest_path", "").strip('"')

        createfolder= self.data.get("new_folder", "").strip('"')
        createname= self.data.get("new_name", "").strip() +".xlsx"
        createfile =os.path.join(createfolder, createname)



        if self.data["use_dest"]:
            from processor import ExcelState
            def worker1():
                    excel_state = ExcelState()
                    excel_state.excel_init(xclfile, headers=headers)
                    
                    for dt, block in dictionary.items():
                        testname = processor.getvaluefromdict(dt, "Test name", block) or "N/A"
                        user     = processor.getvaluefromdict(dt, "Current user", block) or "N/A"

                        for j in range(6):
                            positions = processor.getalarmrowfromdict(block, 1+j, 0)
                            serials   = processor.getalarmrowfromdict(block, 1+j, 1)
                            triggers  = processor.getalarmrowfromdict(block, 1+j, 2)
                            results   = processor.getalarmrowfromdict(block, 1+j, 3)

                            rowdata = {
                                "Date/Time": f"{dt[0]} {dt[1]}",
                                "Test name": testname,
                                "Current user": user,
                                "Position": positions[0] if positions else "N/A",
                                "Serial": serials[0] if serials else "N/A",
                                "Trigger point": triggers[0] if triggers else "N/A",
                                "Results": results[0] if results else "N/A"
                            }

                            t1=self.data["threshold1"]
                            t2=self.data["threshold2"]

                            if not t1:
                                t1=None
                            else:
                                t1 = float(t1)
                            if not t2:
                                t2=None
                            else:
                                t2 = float(t2)    


                            excel_state.excel_write_row(rowdata, t1, t2)


                    sheet = excel_state.sheet


                    serial_index, trigger_index = None, None
                    for col in range(1, sheet.max_column+1):
                        header = sheet.cell(row=1, column=col).value
                        if header == "Serial":
                            serial_index = col
                        elif header == "Trigger point":
                            trigger_index = col

                    if serial_index is None or trigger_index is None:
                        raise Exception("Could not find Serial or Trigger point columns in Excel")

                    serials_clean, triggers_clean = [], []
                    for r in range(2, sheet.max_row+1):
                        serial_val = sheet.cell(row=r, column=serial_index).value
                        trig_val   = sheet.cell(row=r, column=trigger_index).value
                        try:
                            trig_val = float(trig_val)
                        except (TypeError, ValueError):
                            trig_val = None

                        if trig_val is not None:
                            serials_clean.append(serial_val)
                            triggers_clean.append(trig_val)

                    try:
                        excel_state.plot_trigger_points_to_excel(
                            serials_clean,
                            triggers_clean,
                            lower=t1,
                            higher=t2,
                            title="Trigger Points and Serials",
                            anchor_cell=f"B{sheet.max_row+3}"
                        )
                    except Exception as e:
                        error_msg = str(e)
                        dummy = tk.Tk()
                        dummy.withdraw()  # hide window
                        msgbox.showerror(
                            "Plotting Error",
                            f"An error occurred while generating the plot:\n\n{error_msg}\n\n"
                            "Tip: This often happens if the dataset is too large or the figure size is invalid."
                        )
                        dummy.destroy()



                    excel_state.autosize_columns()
                    try:
                        excel_state.excel_save(xclfile)
                        msgbox.showinfo("Success", f"Excel file successfully saved:\n{xclfile}")
                    except Exception as e:
                        msgbox.showerror("Save Error", f"Could not save Excel file:\n{xclfile}\n\nError details:\n{e}")





        if self.data["use_create"]:
            from processor import ExcelState
            def worker2():
                    excel_state = ExcelState()
                    excel_state.excel_init(createfile, headers=headers, create_new=True)
                    
                    for dt, block in dictionary.items():
                        testname = processor.getvaluefromdict(dt, "Test name", block) or "N/A"
                        user     = processor.getvaluefromdict(dt, "Current user", block) or "N/A"

                        for j in range(6):
                            positions = processor.getalarmrowfromdict(block, 1+j, 0)
                            serials   = processor.getalarmrowfromdict(block, 1+j, 1)
                            triggers  = processor.getalarmrowfromdict(block, 1+j, 2)
                            results   = processor.getalarmrowfromdict(block, 1+j, 3)

                            rowdata = {
                                "Date/Time": f"{dt[0]} {dt[1]}",
                                "Test name": testname,
                                "Current user": user,
                                "Position": positions[0] if positions else "N/A",
                                "Serial": serials[0] if serials else "N/A",
                                "Trigger point": triggers[0] if triggers else "N/A",
                                "Results": results[0] if results else "N/A"
                            }
                            t1=self.data["threshold1"]
                            t2=self.data["threshold2"]

                            if not t1:
                                t1=None
                            else:
                                t1 = float(t1)
                            if not t2:
                                t2=None
                            else:
                                t2 = float(t2)    


                            excel_state.excel_write_row(rowdata, t1, t2)


                    sheet = excel_state.sheet


                    serial_index, trigger_index = None, None
                    for col in range(1, sheet.max_column+1):
                        header = sheet.cell(row=1, column=col).value
                        if header == "Serial":
                            serial_index = col
                        elif header == "Trigger point":
                            trigger_index = col

                    if serial_index is None or trigger_index is None:
                        raise Exception("Could not find Serial or Trigger point columns in Excel")

                    serials_clean, triggers_clean = [], []
                    for r in range(2, sheet.max_row+1):
                        serial_val = sheet.cell(row=r, column=serial_index).value
                        trig_val   = sheet.cell(row=r, column=trigger_index).value
                        try:
                            trig_val = float(trig_val)
                        except (TypeError, ValueError):
                            trig_val = None

                        if trig_val is not None:
                            serials_clean.append(serial_val)
                            triggers_clean.append(trig_val)

                    try:
                        excel_state.plot_trigger_points_to_excel(
                            serials_clean,
                            triggers_clean,
                            lower=t1,
                            higher=t2,
                            title="Trigger Points and Serials",
                            anchor_cell=f"B{sheet.max_row+3}"
                        )
                    except Exception as e:
                        error_msg = str(e)
                        dummy = tk.Tk()
                        dummy.withdraw()  # hide window
                        msgbox.showerror(
                            "Plotting Error",
                            f"An error occurred while generating the plot:\n\n{error_msg}\n\n"
                            "Tip: This often happens if the dataset is too large or the figure size is invalid."
                        )
                        dummy.destroy()


                    excel_state.autosize_columns()
                    try:
                        excel_state.excel_save(createfile)
                        msgbox.showinfo("Success", f"Excel file successfully saved:\n{xclfile}")
                    except Exception as e:
                        msgbox.showerror("Save Error", f"Failed to create Excel file:\n{createfile}\n\n{e}")
                    



        if self.data["use_dest"]:
            threading.Thread(target=worker1, daemon=True).start()

        if self.data["use_create"]:
            threading.Thread(target=worker2, daemon=True).start()
            

        




