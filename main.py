import tkinter as tk
import pathlib
from tkinter import filedialog
from tkinter import messagebox
import numpy as np
import pandas as pd
import math
from gmplot import gmplot


class App(tk.Frame):

    def __init__(self, root):
        tk.Frame.__init__(self, root)
        self.root = root
        company_name = "Company"
        self.winfo_toplevel().title(company_name)

        self.make_widgets()
        self.set_position_on_screen()

    def make_widgets(self):
        """
        Create the app widgets
        """

        # font and paddings to make the entries and labels more readable
        self.frame_padding = {"pady": 7, 'padx': 15}
        self.font = "Arial 13"
        self.entry_ipady = 2
        self.entry_width = 10

        self.make_lon_lat_widgets()
        self.make_filedialog_widgets()
        self.make_apply_widgets()

    def make_lon_lat_widgets(self):
        """
        Create the frames that will store the longitude and latiutude entries and labels
        """

        lat_frame = tk.Frame(self.root)
        lat_frame.grid(row=1, column=0, **self.frame_padding)

        lon_frame = tk.Frame(self.root)
        lon_frame.grid(row=2, column=0, **self.frame_padding)

        self.lat_deg = tk.Entry(
            lat_frame, width=self.entry_width, font=self.font)
        self.lat_deg.grid(row=0, column=1, ipady=self.entry_ipady)

        self.lat_min = tk.Entry(
            lat_frame, width=self.entry_width, font=self.font)
        self.lat_min.grid(row=0, column=3, ipady=self.entry_ipady)

        self.lat_sec = tk.Entry(
            lat_frame, width=self.entry_width, font=self.font)
        self.lat_sec.grid(row=0, column=5, ipady=self.entry_ipady)

        self.lon_deg = tk.Entry(
            lon_frame, width=self.entry_width, font=self.font)

        self.lon_deg.grid(row=0, column=1, ipady=self.entry_ipady)

        self.lon_min = tk.Entry(
            lon_frame, width=self.entry_width, font=self.font)
        self.lon_min.grid(row=0, column=3, ipady=self.entry_ipady)

        self.lon_sec = tk.Entry(
            lon_frame, width=self.entry_width, font=self.font)
        self.lon_sec.grid(row=0, column=5, ipady=self.entry_ipady)

        tk.Label(lat_frame, text="Latitude:", width=15,
                 font=self.font).grid(row=0, column=0)

        tk.Label(lat_frame, text="°", font=self.font).grid(row=0, column=2)

        tk.Label(lat_frame, text="'", font=self.font).grid(row=0, column=4)

        tk.Label(lat_frame, text='"N', font=self.font).grid(row=0, column=6)

        tk.Label(lon_frame, text="Longitude:", width=15,
                 font=self.font).grid(row=0, column=0)

        tk.Label(lon_frame, text="°", font=self.font).grid(row=0, column=2)

        tk.Label(lon_frame, text="'", font=self.font).grid(row=0, column=4)

        tk.Label(lon_frame, text='"E', font=self.font).grid(row=0, column=6)

    def make_filedialog_widgets(self):
        """
        Create a frame that will store the filedialog label and button
        """

        file_frame = tk.Frame(self.root)
        file_frame.grid(row=3, column=0, sticky='ew', **self.frame_padding)
        file_frame.columnconfigure(0, weight=3)
        file_frame.columnconfigure(1, weight=1)
        self.home_dir = str(pathlib.Path.home())
        self.file_name = tk.StringVar(self.root, value=self.home_dir)

        self.file_label = tk.Entry(
            file_frame, textvariable=self.file_name, font=self.font, width=35)
        self.file_label.config(state="disabled")
        self.file_label.grid(row=0, column=0, ipady=self.entry_ipady)

        self.choose_button = tk.Button(file_frame, text="Choose file",
                                       font=self.font, command=self.choose_file)
        self.choose_button.grid(row=0, column=1)

    def choose_file(self):
        """
        Open a filedialog and let user to choose an excel file, store its name in the filename attribute
        """

        file_name = filedialog.askopenfilename(initialdir=self.home_dir,
                                               title="Select file", filetypes=[("Excel file", "*.xlsx"), ("Excel file", "*.xls")])
        if file_name:
            self.file_name.set(file_name)

    def make_apply_widgets(self):
        """
        Create a frame with an apply button
        """

        apply_frame = tk.Frame(self.root)
        apply_frame.grid(row=4, column=0, **self.frame_padding)
        self.apply_button = tk.Button(
            apply_frame, text="Apply", font=self.font, command=self.apply)
        self.apply_button.pack()

    def apply(self):
        """
        Read user input and pass it to the method calculating coordinates
        Save the result in excel file
        Create visualization with gmplot package
        """

        # Try to parse strings in entries to float values
        # Show a warning box if parsing raises an exception
        try:
            longitude = (float(self.lon_deg.get())
                         + float(self.lon_min.get()) / 60
                         + float(self.lon_sec.get()) / 3600)

            latiutude = (float(self.lat_deg.get())
                         + float(self.lat_min.get()) / 60
                         + float(self.lat_sec.get()) / 3600)
            if 180 < longitude or longitude < -180 or 90 < latiutude or latiutude < -90:
                raise ValueError

        except ValueError:
            messagebox.showwarning(
                "", "Cordinates must be real numbers!\nLongitude varies from -180 to 180 degrees\nLatitude varies -90 to 90 degrees\nEntries cannot be left empty!")
            return

        try:
            # first two columns should be 'count' and 'name', since they are useless they are skipped
            data_frame = pd.read_excel(self.file_name.get(), usecols=[
                2, 3, 4], index_col=0, dtype={'Position X': np.float64, 'Position Y': np.float64})
            if not set(['Position X', 'Position Y']).issubset(data_frame.columns):
                raise ValueError
        except Exception:
            messagebox.showwarning(
                "", "Invalid excel file")
            return

        coordinates_df = App.calculate_cooridnates(
            data_frame=data_frame, lat=latiutude, lon=longitude)

        path = pathlib.Path(self.file_name.get()).parents[0]

        App.save_data_frame(data_frame=coordinates_df, save_dir=path)
        App.create_visualization(data_frame=coordinates_df, save_dir=path)

        self.end_app()

    @staticmethod
    def create_visualization(data_frame, save_dir, name='visualization.html'):
        """
        Create visualization using gmplot package so user can check if appliaction calculated correct coordinates.
        """

        save_dir = pathlib.PurePath(save_dir, name)
        lat = data_frame.loc[999, 'Latitude Deg']
        lon = data_frame.loc[999, 'Longitude Deg']
        zoom = 17
        gmap = gmplot.GoogleMapPlotter(
            lat, lon, zoom)
        gmap.scatter(data_frame['Latitude Deg'],
                     data_frame['Longitude Deg'], '#FFFF00', size=3, marker=False)

        gmap.draw(save_dir)

    @staticmethod
    def save_data_frame(data_frame, save_dir, name='coordinates.xls'):
        """
        Save data frame in excel format in given directory
        """

        save_dir = pathlib.PurePath(save_dir, name)
        data_frame.to_excel(save_dir, columns=['Coordinates'])

    @staticmethod
    def calculate_cooridnates(data_frame, lat, lon):
        """
        Calucalte geographic coridnates of points given in data frame
        knowing geographic coordinates of starting point.
        Return data given with columns representing geographic coordinates in DD and DMS format
        """

        starting_point_index = 999
        starting_point = data_frame.loc[starting_point_index]
        data_frame['Distance X'] = (data_frame['Position X']
                                    - starting_point['Position X'])
        data_frame['Distance Y'] = (data_frame['Position Y']
                                    - starting_point['Position Y'])
        # units used in the data frame are centimeters so distances are divided by 1000 to be in meters
        scale = 1000
        data_frame['Distance X'] /= scale
        data_frame['Distance Y'] /= scale

        # https://en.wikipedia.org/wiki/Decimal_degrees
        # 'each degree at the equator represents 111,319.9 meters'
        # 'however, one degree of longitude is multiplied by the cosine of the latitude'
        ratio = 111319.9
        data_frame['Latitude Deg'] = data_frame['Distance Y'] / ratio + lat

        data_frame['Longitude Deg'] = data_frame.apply(lambda row: row['Distance X'] / (
            ratio * math.cos(math.radians(row['Latitude Deg']))) + lon, axis=1)
        data_frame['Latitude DMS'] = data_frame.apply(lambda row: str(math.floor(row['Latitude Deg'])) + '°'
                                                      + str(math.floor(row['Latitude Deg'] * 60 % 60)) + "'"
                                                      + str(math.floor(row['Latitude Deg'] * 3600 % 60)) + '.'
                                                      + str(math.floor(row['Latitude Deg'] * 360000) % 100) + '"N', axis=1)
        data_frame['Longitude DMS'] = data_frame.apply(lambda row: str(math.floor(row['Longitude Deg'])) + '°'
                                                       + str(math.floor(row['Longitude Deg'] * 60 % 60)) + "'"
                                                       + str(math.floor(row['Longitude Deg'] * 3600 % 60)) + '.'
                                                       + str(math.floor(row['Longitude Deg'] * 360000) % 100) + '"E', axis=1)
        data_frame['Coordinates'] = (data_frame['Latitude DMS']
                                     + data_frame['Longitude DMS'])
        data_frame = data_frame.sort_index()

        return data_frame

    def end_app(self):
        messagebox.showinfo(
            "Succes",
            "Coordinates calculated"
        )
        self.root.destroy()

    def set_position_on_screen(self):
        """
        Center the app window on the screen
        """

        self.winfo_toplevel().update_idletasks()
        window_height = self.root.winfo_reqheight()
        window_width = self.winfo_toplevel().winfo_reqwidth()
        screen_height = self.root.winfo_screenheight()
        screen_width = self.root.winfo_screenwidth()

        center_height = int(screen_height / 2 - window_height / 2)
        center_width = int(screen_width / 2 - window_width / 2)

        self.root.geometry("+{}+{}".format(center_width, center_height))


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
