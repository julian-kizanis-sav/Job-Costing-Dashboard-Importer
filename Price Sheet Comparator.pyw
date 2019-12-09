# SAV Digital Environments
# Julian Kizanis
# generated in part by wxGlade 0.9.4 on Mon Nov 18 07:49:50 2019
#

from datetime import date
from getpass import getuser
from ntpath import basename

import wx
from openpyxl import load_workbook
import pandas as pd
from fuzzywuzzy import fuzz
import time
import threading
import multiprocessing

# Declare GUI Constants
MENU_FILE_EXIT = wx.ID_ANY
DRAG_SOURCE = wx.ID_ANY

# Global Variables
pb = False

# Global Constants
PRICE_UPDATE = 1
DISCONTINUED_SEARCH = 2
CONTINUE = 2
OVERRIDE = -2
CANCEL = -1


def new_window_loop(local_import_files, compare_type, compare_list, import_list):
    new_window = ComparisonProcess(0)
    compare_thread = threading.Thread(target=sheet_compare, args=(new_window.frame, local_import_files,
                                                                  compare_type,
                                                                  compare_list,
                                                                  import_list),
                                      daemon=True)
    compare_thread.start()

    new_window.MainLoop()


def sheet_compare(daemon_window, local_import_files, compare_type, individual_list, master_list):
    manufacturers = []
    pandas_matches = panda_match_ini()
    pandas_perfect_matches = panda_match_ini()
    pandas_uncertain_matches = panda_match_ini()
    index_matches = 0
    pandas_dtools = None
    # still_running = RunningComparisonsDialog(None, wx.ID_ANY, "")
    # still_running.Show()
    start_time = time.time()
    print(local_import_files)
    for file in local_import_files:
        if file is local_import_files[0]:
            try:
                pandas_dtools = pd.read_csv(file)
                continue
            except UnicodeDecodeError:
                wx.MessageBox("Please save your dtools 'CSV (Comma delimited) (*.csv)' file "
                              "as a 'CSV UTF-8 (Comma delimited) (*.csv)' file.", "Error", wx.OK | wx.ICON_INFORMATION)
                return False
        try:
            pandas_vendor = pd.read_csv(file)
        except UnicodeDecodeError:
            wx.MessageBox("Please save your 'CSV (Comma delimited) (*.csv)' file as "
                          "a 'CSV UTF-8 (Comma delimited) (*.csv)' file.", "Error", wx.OK | wx.ICON_INFORMATION)
            continue
        if compare_type == PRICE_UPDATE:
            for index_dtools, part_dtools in enumerate(pandas_dtools.loc[:, 'Model']):
                pandas_matches = match_new_row(pandas_matches, pandas_dtools, index_dtools)
                for index_vendor, part_vendor in enumerate(pandas_vendor.loc[:, 'Model']):
                    if not (pandas_vendor.loc[index_vendor, 'Manufacturer'] and
                            pandas_vendor.loc[index_vendor, 'Model'] and
                            pandas_vendor.loc[index_vendor, 'Unit Cost'] and
                            pandas_vendor.loc[index_vendor, 'Unit Price']):
                        continue
                    if not manufacturers:
                        manufacturers.append(pandas_vendor.loc[index_vendor, 'Manufacturer'])
                    pro_dtool_part = str(part_dtools).lower().strip(" -_(')/")
                    pro_vendor_part = str(part_vendor).lower().strip(" -_(')/")
                    temp_match_ratio = (fuzz.partial_ratio(pro_dtool_part, pro_vendor_part) +
                                        fuzz.ratio(pro_dtool_part, pro_vendor_part) +
                                        fuzz.token_sort_ratio(part_dtools, part_vendor)) / 3
                    if temp_match_ratio > pandas_matches.loc[index_dtools, 'Match Ratio']:
                        new_best_match(pandas_matches, index_dtools, temp_match_ratio, pandas_vendor, index_vendor)

                    daemon_window.list_ctrl_1.Append((pandas_matches.loc[index_matches, 'Manufacturer'], part_dtools,
                                                      pandas_matches.loc[index_matches, 'Match Ratio'],
                                                      time.time() - start_time, index_matches))
                    daemon_window.list_ctrl_1.EnsureVisible(daemon_window.list_ctrl_1.GetItemCount() - 1)
                    daemon_window.Update()
                print(part_dtools, index_dtools)
            pandas_matches = pandas_matches.sort_values('Match Ratio', ascending=False)
        elif compare_type == DISCONTINUED_SEARCH:
            # manufacturers = []
            index_manufacturers = 0
            for manufacturer_vendor in pandas_vendor.loc[:, 'Manufacturer']:
                if not manufacturers:
                    manufacturers.append(manufacturer_vendor)
                elif manufacturer_vendor != manufacturers[index_manufacturers]:
                    index_manufacturers += 1
                    manufacturers.append(manufacturer_vendor)
            for manufacturer in manufacturers:
                for index_dtools, part_dtools in enumerate(pandas_dtools.loc[:, 'Model']):
                    number_dtools = pandas_dtools.loc[index_dtools, 'Part Number']
                    if manufacturer.lower() == pandas_dtools.loc[index_dtools, 'Manufacturer'].lower():
                        pandas_matches = match_new_row(pandas_matches, pandas_dtools, index_dtools)
                        for index_vendor, part_vendor in enumerate(pandas_vendor.loc[:, 'Model']):
                            if not (pandas_vendor.loc[index_vendor, 'Manufacturer'] and
                                    pandas_vendor.loc[index_vendor, 'Model'] and
                                    pandas_vendor.loc[index_vendor, 'Unit Cost'] and
                                    pandas_vendor.loc[index_vendor, 'Unit Price']):
                                continue
                            pro_dtool_part = str(part_dtools).lower().strip(" -_(')/")
                            pro_vendor_part = str(part_vendor).lower().strip(" -_(')/")

                            temp_match_ratio = (fuzz.partial_ratio(pro_dtool_part, pro_vendor_part) +
                                                fuzz.ratio(pro_dtool_part, pro_vendor_part) +
                                                fuzz.token_sort_ratio(part_dtools, part_vendor)) / 3
                            if temp_match_ratio > pandas_matches.loc[index_matches, 'Match Ratio']:
                                new_best_match(pandas_matches, index_matches, temp_match_ratio, pandas_vendor,
                                               index_vendor)
                            if number_dtools:
                                pro_dtool_number = str(number_dtools).lower().strip(" -_(')/")
                                temp_match_ratio = (fuzz.partial_ratio(pro_dtool_number, pro_vendor_part) +
                                                    fuzz.ratio(pro_dtool_number, pro_vendor_part) +
                                                    fuzz.token_sort_ratio(number_dtools, part_vendor)) / 3
                                if temp_match_ratio > pandas_matches.loc[index_matches, 'Match Ratio']:
                                    new_best_match(pandas_matches, index_matches, temp_match_ratio, pandas_vendor,
                                                   index_vendor)

                        # print(part_dtools, index_matches)
                        daemon_window.list_ctrl_1.Append((manufacturer, part_dtools,
                                                          pandas_matches.loc[index_matches, 'Match Ratio'],
                                                          time.time() - start_time, index_matches))
                        daemon_window.list_ctrl_1.EnsureVisible(daemon_window.list_ctrl_1.GetItemCount() - 1)
                        daemon_window.Update()
                        index_matches += 1

                for index_matches, match_ratio in enumerate(pandas_matches.loc[:, 'Match Ratio']):
                    if match_ratio == 100:
                        pandas_perfect_matches = pandas_perfect_matches.append(pandas_matches.loc[index_matches, :])
                    else:
                        pandas_uncertain_matches = pandas_uncertain_matches.append(pandas_matches.loc[index_matches, :])
                pandas_matches = pandas_matches.sort_values('Match Ratio', ascending=False)
                pandas_uncertain_matches = pandas_uncertain_matches.sort_values('Match Ratio', ascending=False)

        else:
            print("Error: ")
            return False
    # pandas_matches = pandas_matches.sort_values('Match Ratio', ascending=False)
    if individual_list:
        print(manufacturers[0])
        pandas_matches.to_csv(f"Comparator Output/Dealer Import Compare Sheet ({manufacturers[0]}) {date.today()}.csv",
                              index=False)
        print("saved")
    if master_list:
        try:
            pandas_master = pd.read_csv("Comparator Output/__Master Dealer Import Compare Sheet.csv")
            pandas_master_perfect = pd.read_csv("Comparator Output/__Perfect Master Dealer Import Compare Sheet.csv")
            pandas_master_uncertain = \
                pd.read_csv("Comparator Output/__Uncertain Master Dealer Import Compare Sheet.csv")
        except FileNotFoundError:
            return False
        pandas_master = pd.concat([pandas_master, pandas_matches], axis=0, sort=False)
        pandas_master_perfect = pd.concat([pandas_master_perfect, pandas_perfect_matches], axis=0, sort=False)
        pandas_master_uncertain = pd.concat([pandas_master_uncertain, pandas_uncertain_matches], axis=0, sort=False)
        # pandas_master = pandas_master.sort_values('Match Ratio', ascending=False)
        pandas_master.to_csv("Comparator Output/__Master Dealer Import Compare Sheet.csv", index=False)
        pandas_master_perfect.to_csv("Comparator Output/__Perfect Master Dealer Import Compare Sheet.csv", index=False)
        pandas_master_uncertain.to_csv("Comparator Output/__Uncertain Master Dealer Import Compare Sheet.csv",
                                       index=False)
    wx.MessageBox("Done!", "Done!", wx.OK | wx.ICON_INFORMATION)
    return pandas_matches


def panda_match_ini():
    match = pd.DataFrame(columns={'Index_dtools', 'Manufacturer_dtools', 'Model_dtools',
                                  'Part_dtools', 'Cost_dtools', 'Price_dtools', 'Used_dtools',
                                  'Match Ratio', 'Index[1]', 'Manufacturer[1]',
                                  'Model[1]', 'Part[1]', 'Cost[1]', 'Price[1]', 'Used[1]', 'Match', 'Keep'})

    return match[['Index_dtools', 'Manufacturer_dtools', 'Model_dtools',
                  'Part_dtools', 'Cost_dtools', 'Price_dtools', 'Used_dtools',
                  'Match Ratio', 'Index[1]', 'Manufacturer[1]',
                  'Model[1]', 'Part[1]', 'Cost[1]', 'Price[1]', 'Used[1]', 'Match', 'Keep']]


def match_new_row(pandas_match, pandas_import, index):
    return pandas_match.append({'Index_dtools': index, 'Manufacturer_dtools': pandas_import.loc[index, "Manufacturer"],
                                'Model_dtools': pandas_import.loc[index, "Model"],
                                'Part_dtools': pandas_import.loc[index, "Part Number"],
                                'Cost_dtools': pandas_import.loc[index, "Unit Cost"],
                                'Price_dtools': pandas_import.loc[index, "Unit Price"],
                                'Used_dtools': False, 'Match Ratio': 0, 'Index[1]': 0, 'Manufacturer[1]': "",
                                'Model[1]': "", 'Part[1]': "", 'Cost[1]': 0, 'Price[1]': 0, 'Used[1]': False,
                                'Match': False, 'Keep': True},
                               ignore_index=True)


def new_best_match(pandas_match, index_match, match_ratio, pandas_import, index_import):
    pandas_match.loc[index_match, "Index[1]"] = index_import
    pandas_match.loc[index_match, "Manufacturer[1]"] = pandas_import.loc[index_import, "Manufacturer"]
    pandas_match.loc[index_match, "Model[1]"] = pandas_import.loc[index_import, "Model"]
    try:
        pandas_match.loc[index_match, "Part[1]"] = pandas_import.loc[index_import, "Part Number"]
    except KeyError:
        pass
    pandas_match.loc[index_match, "Cost[1]"] = pandas_import.loc[index_import, "Unit Cost"]
    pandas_match.loc[index_match, "Price[1]"] = pandas_import.loc[index_import, "Unit Price"]
    pandas_match.loc[index_match, "Match Ratio"] = match_ratio


def check_for_duplicates(import_directory, imported_list):
    """This function checks if the file has already been imported"""
    for imp in imported_list:  # cycles through the directories of the previously imported files
        if import_directory == imp:  # if the directory we are trying to import matches a directory in the database
            return True  # we found a match
    return False  # no matches were found


def open_spreadsheet(directory):
    """This function tries to open a spreadsheet and prompts the user if it cannot"""
    while True:  # infinite loop
        try:
            dashboard = load_workbook(filename=directory, read_only=False, data_only=True)  # tries to open spreadsheet
            return dashboard  # returns the spreadsheet if it was opened
        except PermissionError:  # the spreadsheet is already open by something else
            dialog = DatasheetOpenDialog(basename(directory), None, wx.ID_ANY, "")  # asks if the user wants to retry
            retry = dialog.ShowModal()  # captures the users response
            if not retry:  # if the user doesn't want to retry
                return None


class FileDropTarget(wx.FileDropTarget):
    """ This object implements Drop Target functionality for Files """

    def __init__(self, text_ctrl, import_files):
        """ Initialize the Drop Target, passing in the Object Reference to
        indicate what should receive the dropped files """
        # Initialize the wxFileDropTarget Object
        wx.FileDropTarget.__init__(self)
        # Store the Object Reference for dropped files
        self.text_ctrl = text_ctrl
        self.import_files = import_files

    def OnDropFiles(self, x, y, file_names):
        """ Implement File Drop """
        # Move Insertion Point to the end of the widget's text
        self.text_ctrl.SetInsertionPointEnd()
        # append a list of the file names dropped
        dup_check = False
        for file in file_names:
            for iFile in self.import_files:
                if file == iFile:
                    dup_check = True
            if not file.endswith('.csv'):
                wx.MessageBox("Incorrect file type. Please choose an .csv file.", "Error", wx.OK | wx.ICON_INFORMATION)
                continue
            if not dup_check:
                self.text_ctrl.WriteText(basename(file) + '\n')
                self.import_files.append(file)
            else:
                print("Removed duplicate import file!")
                wx.MessageBox("File already in import list.", "Error", wx.OK | wx.ICON_INFORMATION)
                dup_check = False
        self.text_ctrl.WriteText('\n')


class FirstFrame(wx.Frame):
    """This object is the main window"""

    def __init__(self, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)

        self.compare_process = []
        self.import_files = []

        self.SetSize((640, 428))
        self.button_browse = wx.FilePickerCtrl(self)
        # self.button_4.Bind(wx.EVT_FILEPICKER_CHANGED, self.on_choose_file)
        self.text_ctrl_drag_drop = wx.TextCtrl(self, wx.ID_ANY, "", style=wx.TE_MULTILINE | wx.TE_READONLY)
        drop_target = FileDropTarget(self.text_ctrl_drag_drop, self.import_files)
        # Link the Drop Target Object to the Text Control
        self.text_ctrl_drag_drop.SetDropTarget(drop_target)

        # initializes the buttons
        self.choice_compare_type = wx.Choice(self, wx.ID_ANY, choices=["Choose Compare Type",
                                                                       "Price Update", "Discontinued Search"])
        self.checkbox_individual_list = wx.CheckBox(self, wx.ID_ANY, "Individual List")
        self.checkbox_importable_list = wx.CheckBox(self, wx.ID_ANY, "Master List")
        # self.checkbox_jake_dashboard = wx.CheckBox(self, wx.ID_ANY, "Jake's Master Dashboard")
        self.panel_1 = wx.Panel(self, wx.ID_ANY)
        self.button_continue = wx.Button(self, wx.ID_ANY, "Continue")
        self.button_cancel = wx.Button(self, wx.ID_ANY, "Cancel")
        self.button_clear = wx.Button(self, wx.ID_ANY, "Clear")

        self.__set_properties()
        self.__do_layout()
        self.SetMinSize((345, 345))

        # initializes the events
        self.Bind(wx.EVT_FILEPICKER_CHANGED, self.on_choose_file, self.button_browse)
        # self.Bind(wx.EVT_CHOICE, self.on_phase_selection, self.choice_compare_type)
        # self.Bind(wx.EVT_CHECKBOX, self.on_general_master_dashboard_checkbox, self.checkbox_comparator_list)
        # self.Bind(wx.EVT_CHECKBOX, self.on_kaceys_master_dashboard_checkbox, self.checkbox_importable_list)
        # self.Bind(wx.EVT_CHECKBOX, self.on_jakes_master_dashboard_checkbox, self.checkbox_jake_dashboard)
        self.Bind(wx.EVT_BUTTON, self.on_continue_from_main_window, self.button_continue)
        self.Bind(wx.EVT_BUTTON, self.on_cancel_program, self.button_cancel)
        self.Bind(wx.EVT_BUTTON, self.on_clear, self.button_clear)
        self.Bind(wx.EVT_ICONIZE, self.on_minimize)

    def __set_properties(self):
        self.SetTitle("Price Sheet Comparator")
        _icon = wx.NullIcon
        _icon.CopyFromBitmap(wx.Bitmap("SAV-Social-Profile.jpg", wx.BITMAP_TYPE_ANY))
        self.SetIcon(_icon)

        self.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.choice_compare_type.SetMinSize((150, 23))
        self.choice_compare_type.SetSelection(0)
        self.checkbox_individual_list.SetValue(1)
        self.checkbox_importable_list.SetValue(1)
        # self.checkbox_jake_dashboard.SetValue(0)

    def __do_layout(self):
        sizer_5 = wx.BoxSizer(wx.VERTICAL)
        sizer_9 = wx.GridBagSizer(0, 0)
        sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8 = wx.BoxSizer(wx.VERTICAL)
        sizer_11 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_15 = wx.BoxSizer(wx.VERTICAL)
        sizer_12 = wx.BoxSizer(wx.VERTICAL)
        sizer_13 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_7 = wx.BoxSizer(wx.VERTICAL)
        sizer_14 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_16 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_10 = wx.BoxSizer(wx.VERTICAL)
        user = getuser()
        label_1 = wx.StaticText(self, wx.ID_ANY, f"Hello {user}! This program compares dealer spreadsheets to a d-tools"
                                                 f" csv file")
        sizer_10.Add(label_1, 0, wx.ALL, 5)
        static_line_1 = wx.StaticLine(self, wx.ID_ANY)
        sizer_10.Add(static_line_1, 0, wx.EXPAND, 0)
        sizer_5.Add(sizer_10, 0, wx.EXPAND, 0)
        sizer_16.Add(self.button_browse, 0, wx.ALL, 12)
        label_6 = wx.StaticText(self, wx.ID_ANY, "Or drag and drop files below")
        sizer_16.Add(label_6, 0, wx.ALIGN_CENTER, 0)
        sizer_7.Add(sizer_16, 0, wx.EXPAND, 0)
        sizer_14.Add(self.text_ctrl_drag_drop, 1, wx.EXPAND, 0)
        sizer_7.Add(sizer_14, 1, wx.EXPAND, 0)
        sizer_6.Add(sizer_7, 2, wx.EXPAND, 0)
        bitmap_2 = wx.StaticBitmap(self, wx.ID_ANY, wx.Bitmap("SAV-Company-Logo.png", wx.BITMAP_TYPE_ANY))
        sizer_12.Add(bitmap_2, 0, wx.BOTTOM | wx.RIGHT | wx.TOP, 12)
        sizer_13.Add(self.choice_compare_type, 0, wx.BOTTOM | wx.LEFT, 6)
        sizer_12.Add(sizer_13, 1, wx.EXPAND, 0)
        sizer_8.Add(sizer_12, 0, wx.EXPAND, 0)
        sizer_15.Add(self.checkbox_individual_list, 0, wx.LEFT | wx.RIGHT | wx.TOP, 6)
        sizer_15.Add(self.checkbox_importable_list, 0, wx.LEFT | wx.RIGHT | wx.TOP, 6)
        # sizer_15.Add(self.checkbox_jake_dashboard, 0, wx.LEFT | wx.RIGHT | wx.TOP, 6)
        sizer_11.Add(sizer_15, 1, wx.EXPAND, 0)
        sizer_8.Add(sizer_11, 1, wx.EXPAND, 0)
        sizer_6.Add(sizer_8, 0, wx.EXPAND | wx.LEFT, 6)
        sizer_5.Add(sizer_6, 1, wx.EXPAND, 0)
        sizer_9.Add(self.panel_1, (0, 0), (1, 1), wx.EXPAND, 0)
        sizer_9.Add(self.button_continue, (0, 1), (1, 1), wx.ALIGN_BOTTOM | wx.ALIGN_RIGHT | wx.ALL, 6)
        sizer_9.Add(self.button_cancel, (0, 3), (1, 1), wx.ALIGN_BOTTOM | wx.ALIGN_RIGHT | wx.ALL, 6)
        sizer_9.Add(self.button_clear, (0, 2), (1, 1), wx.ALIGN_BOTTOM | wx.ALIGN_RIGHT | wx.ALL, 6)
        sizer_5.Add(sizer_9, 0, wx.ALIGN_BOTTOM | wx.ALIGN_RIGHT | wx.ALL | wx.EXPAND, 12)
        self.SetSizer(sizer_5)
        self.Layout()

    def on_choose_file(self, event):  # button_browse
        dup_check = False
        file = self.button_browse.GetPath()  # catches what file the user chose
        for iFile in self.import_files:  # checks if file is already in the to-be imported list
            if file == iFile:
                dup_check = True
        if not file.endswith('.csv'):
            wx.MessageBox("Incorrect file type. Please choose .csv file.", "Error", wx.OK | wx.ICON_INFORMATION)
            event.skip()
        if not dup_check:
            self.import_files.append(file)
            self.text_ctrl_drag_drop.WriteText(basename(file) + '\n')  # shows the user they chose this
        else:
            print("Removed duplicate import file!")
            wx.MessageBox("File already in import list.", "Error", wx.OK | wx.ICON_INFORMATION)
        event.Skip()

    def on_continue_from_main_window(self, event):  # event handler
        print(self.import_files)
        dtools_date = date.today().strftime("Products %b %d, %Y")
        dtools_present = False
        if len(self.import_files) < 2:
            wx.MessageBox("You need at least 2 files to continue.", "Error", wx.OK | wx.ICON_INFORMATION)
            return False
        if self.choice_compare_type.GetSelection() == 0:
            wx.MessageBox("You need to choose the comparison type.", "Error", wx.OK | wx.ICON_INFORMATION)
            return False

        for index, file in enumerate(self.import_files):
            print(basename(dtools_date), file)
            if basename(file).startswith(dtools_date):
                temp_file = self.import_files[0]
                self.import_files[0] = file
                self.import_files[index] = temp_file
                dtools_present = True
                break
        if not dtools_present:
            wx.MessageBox("You need to import a current dtools spreadsheet. "
                          "Please do not change the name the file name", "Error", wx.OK | wx.ICON_INFORMATION)
            return False

        # daemon_window = threading.Thread(target=new_window_loop, args=(), daemon=True)
        # daemon_window.start()
        # if len(self.import_files == 2):
        #     self.compare_process.append(multiprocessing.Process(target=new_window_loop,
        #                                                         args=(self.import_files,
        #                                                               self.choice_compare_type.GetSelection(),
        #                                                               self.checkbox_comparator_list.GetValue(),
        #                                                               self.checkbox_importable_list.GetValue()),
        #                                                         daemon=True))
        #     self.compare_process[len(self.compare_process) - 1].start()
        else:
            for file in self.import_files[1:]:
                temp_import_files = (self.import_files[0], file)
                self.compare_process.append(multiprocessing.Process(target=new_window_loop,
                                                                    args=(temp_import_files,
                                                                          self.choice_compare_type.GetSelection(),
                                                                          self.checkbox_individual_list.GetValue(),
                                                                          self.checkbox_importable_list.GetValue()),
                                                                    daemon=True))
                self.compare_process[len(self.compare_process) - 1].start()
        # sheet_compare(self.choice_compare_type.GetSelection(), self.checkbox_comparator_list.GetValue(),
        #               self.checkbox_importable_list.GetValue())
        self.text_ctrl_drag_drop.SetValue("")  # resets the program
        self.import_files.clear()  # resets the program
        event.Skip()

    def on_cancel_program(self, event):  # event handler
        print(getuser())
        self.Destroy()
        event.Skip()

    def on_clear(self, event):  # resets the program
        self.text_ctrl_drag_drop.SetValue("")
        global pb
        self.import_files.clear()
        pb = not pb
        event.Skip()

    @staticmethod
    def on_minimize(event):  # easter egg
        global pb
        if pb:
            wx.MessageBox("Or is it Peanutbutter?", "Peanut butter!", wx.OK | wx.ICON_INFORMATION)
        pb = False
        event.Skip()


class RunningComparisonsFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP
        wx.Frame.__init__(self, *args, **kwds)
        self.list_ctrl_1 = wx.ListCtrl(self, wx.ID_ANY, style=wx.LC_HRULES | wx.LC_REPORT | wx.LC_VRULES)

        self.__set_properties()
        self.__do_layout()

    def __set_properties(self):
        self.SetTitle("Running Comparisons")
        self.list_ctrl_1.AppendColumn("Manufacturer", format=wx.LIST_FORMAT_LEFT, width=-1)
        self.list_ctrl_1.AppendColumn("Model Number", format=wx.LIST_FORMAT_LEFT, width=-1)
        self.list_ctrl_1.AppendColumn("Match %", format=wx.LIST_FORMAT_LEFT, width=-1)
        self.list_ctrl_1.AppendColumn("Elapsed Time", format=wx.LIST_FORMAT_LEFT, width=-1)
        self.list_ctrl_1.AppendColumn("Index", format=wx.LIST_FORMAT_LEFT, width=-1)

    def __do_layout(self):
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        runnig_label = wx.StaticText(self, wx.ID_ANY, "Crunching the numbers...")
        sizer_1.Add(runnig_label, 0, 0, 0)
        sizer_1.Add(self.list_ctrl_1, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        sizer_1.Fit(self)
        self.Layout()

    def on_cancel_program(self, event):  # event handler
        print(getuser())
        self.Destroy()
        event.Skip()


class DatasheetOpenDialog(wx.Dialog):
    def __init__(self, open_data_sheet, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_DIALOG_STYLE
        wx.Dialog.__init__(self, *args, **kwds)
        self.open_data_sheet = open_data_sheet
        self.text_ctrl_open_datasheet = wx.TextCtrl(self, wx.ID_ANY,
                                                    f"{open_data_sheet} is open by a user. Please close "
                                                    f"{open_data_sheet} and click \"Retry\".",
                                                    style=wx.BORDER_NONE | wx.TE_MULTILINE | wx.TE_READONLY)
        self.panel_2 = wx.Panel(self, wx.ID_ANY)
        self.button_1 = wx.Button(self, wx.ID_ANY, "Retry")
        self.button_5 = wx.Button(self, wx.ID_ANY, "Back")

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_TEXT, self.text_ctrl_open_data_sheet, self.text_ctrl_open_datasheet)
        self.Bind(wx.EVT_BUTTON, self.on_retry, self.button_1)
        self.Bind(wx.EVT_BUTTON, self.on_back, self.button_5)

    def __set_properties(self):
        _icon = wx.NullIcon
        _icon.CopyFromBitmap(wx.Bitmap("SAV-Social-Profile.jpg", wx.BITMAP_TYPE_ANY))
        self.SetIcon(_icon)

        self.SetTitle("dialog_3")
        self.text_ctrl_open_datasheet.SetBackgroundColour(wx.Colour(255, 255, 255))

    def __do_layout(self):
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(self.text_ctrl_open_datasheet, 0, wx.ALL | wx.EXPAND, 12)
        sizer_2.Add(self.panel_2, 1, 0, 0)
        sizer_2.Add(self.button_1, 0, wx.ALIGN_BOTTOM | wx.ALL | wx.FIXED_MINSIZE, 12)
        sizer_2.Add(self.button_5, 0, wx.ALIGN_BOTTOM | wx.ALL | wx.FIXED_MINSIZE, 12)
        sizer_1.Add(sizer_2, 1, wx.ALIGN_BOTTOM | wx.ALIGN_RIGHT | wx.ALL | wx.EXPAND | wx.FIXED_MINSIZE, 1)
        self.SetSizer(sizer_1)
        sizer_1.Fit(self)
        self.Layout()

    def text_ctrl_open_data_sheet(self, event):  # event handler
        print(f"{self.open_data_sheet} is currently open by a user!")
        event.Skip()

    def on_retry(self, event):  # event handler
        self.EndModal(True)
        self.Destroy()
        event.Skip()

    def on_back(self, event):  # event handler
        self.EndModal(False)
        self.Destroy()
        event.Skip()


class MyApp(wx.App):
    def OnInit(self):
        self.frame = FirstFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


class ComparisonProcess(wx.App):
    def OnInit(self):
        self.frame = RunningComparisonsFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


if __name__ == "__main__":
    ImportProjectDatasheets = MyApp(0)
    ImportProjectDatasheets.MainLoop()
