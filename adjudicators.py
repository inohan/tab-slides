import pandas as pd
from pptx import Presentation
from out_ppt import create_slides_adjudicators, BlankSlideSettings
from pathlib import Path
import inquirer
from tkinter import filedialog
from collections import defaultdict
from exception import MissingColumnError, LogoError, enc
import logging

class AdjudicatorSlidesBuilder():
    def __init__(self):
        self.df_participants = None
        self.df_standings = None
        self.dict_logos = defaultdict(None)
        self.dict_layouts = defaultdict(int)
        self.pres = None
        self.paths = {
            "logo": None,
            "participant": None,
            "standing": None,
            "presentation": None
        }
        self.display = "avg. {} points"
        self.txt_title = "{} Best Adjudicator"
        self.danger_prevention = BlankSlideSettings.CHANGE_RANK
        try:
            self.load_logos("./institution_logo.csv")
            self.paths["logo"] = "./institution_logo.csv"
        except Exception as e:
            print(e)

    def main_menu(self):
        txt_participant = f"Participant: {self.paths['participant']}" if self.paths["participant"] else "Participant: NOT LOADED"
        txt_logo = f"Logo: {self.paths['logo']}" if self.paths["logo"] else "Logo: NOT LOADED"
        txt_standing = f"Standing: {self.paths['standing']}" if self.paths["standing"] else "Standing: NOT LOADED"
        txt_presentation = f"Presentation: {self.paths['presentation']}" if self.paths["presentation"] else "Presentation: NOT LOADED"
        txt_title = f"Title: \"{self.txt_title.format('(nth)')}\""
        txt_display = f"Point display: \"{self.display}\""
        choices = [
            (txt_participant, "participant"),
            (txt_logo, "logo"),
            (txt_standing, "standing"),
            (txt_presentation, "presentation"),
            (txt_title, "title"),
            (txt_display, "display"),
            (f"Danger prevention: {self.danger_prevention.value}", "danger"),
            ("Create slide", "create"),
            ("Quit", "quit")
        ]
        selected = inquirer.list_input("Configure", choices=choices)
        if selected == "participant":
            self.prompt_path_participant()
        elif selected == "logo":
            self.prompt_path_logo()
        elif selected == "standing":
            self.prompt_path_standing()
        elif selected == "presentation":
            self.prompt_path_presentation()
        elif selected == "title":
            self.prompt_title()
        elif selected == "display":
            self.prompt_metric_display()
        elif selected == "danger":
            self.prompt_danger_prevention()
        elif selected == "create":
            self.create_slides()
        
        if selected != "quit":
            self.main_menu()
    
    def prompt_path_logo(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select institution-logo file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.load_logos(file_path)
                self.paths["logo"] = file_path
            except Exception as e:
                print(e)
        else:
            print("Error: file does not exist.")

    def prompt_path_participant(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select adjudicator participant file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.load_participants(file_path)
                self.paths["participant"] = file_path
            except Exception as e:
                print(e)
        else:
            print("Error: file does not exist.")

    def prompt_path_standing(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select adjudicator standing file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.load_standings(file_path)
                self.paths["standing"] = file_path
            except Exception as e:
                logging.exception("What?")
        else:
            print("Error: file does not exist.")
    
    def prompt_path_presentation(self):
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint file", "*.pptx")], title="Select PowerPoint")
        if len(file_path) and Path(file_path).exists():
            try:
                self.load_presentation(file_path)
                self.paths["presentation"] = file_path
            except Exception as e:
                print(e)
        else:
            print("Error: file does not exist.")

    def prompt_title(self):
        self.txt_title = inquirer.text("Title text (e.g. \"(nth) Best Adjudicator\")", default="(nth) Best Adjudicator").replace("(nth)", "{}")

    def prompt_metric_display(self):
        self.display = inquirer.text("Enter format for displaying scores (e.g. \"avg. {{}} points\")")

    def prompt_danger_prevention(self):
        choices = [(e.value, e) for e in BlankSlideSettings]
        self.danger_prevention = inquirer.list_input("Select danger prevention settings", default=self.danger_prevention, choices = choices)

    def load_logos(self, path):
        df_logo = pd.read_csv(path, encoding = enc)
        # Check for missing columns
        missing_columns = [item for item in ["institution", "path"] if item not in df_logo.columns]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {path}: {', '.join(missing_columns)}")
        # Update dictionary
        for index, row in df_logo.iterrows():
            self.dict_logos[row["institution"]] = str(Path(row["path"])) if Path(row["path"]).exists() else None

    def load_participants(self, path):
        self.df_participants = pd.read_csv(path, encoding = enc)
        # Check for missing columns
        missing_columns = [item for item in ["name", "institution"] if item not in self.df_participants.columns]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {path}: {', '.join(missing_columns)}")

    def load_standings(self, path):
        df_load = pd.read_csv(path, encoding = enc)
        print(df_load.columns)
        missing_columns = [item for item in ["name", "score"] if item not in df_load]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {path}: {', '.join(missing_columns)}")
        # Selecting ranks
        col_rank = inquirer.list_input("Which column holds the rank?", choices = [c for c in df_load.columns if c not in ["name", "score"]])
        # Selecting show metrics
        df_filter = df_load[["name", "score", col_rank]].rename(columns={col_rank: "rank"}).dropna(subset=["rank"]).sort_values(by=["rank", "name"], ascending=[False, True])
        self.df_standings = df_filter

    def load_presentation(self, path):
        # Reset
        self.dict_layouts = defaultdict(int)
        self.pres = Presentation(path)
        choices = [(f"[{index}] {layout.name}", index) for index, layout in enumerate(self.pres.slide_layouts)]
        for i in range(2):
            self.dict_layouts[i] = inquirer.list_input(f"Select slide for {i} institutions", choices=choices)

    def create_slides(self):
        if not (self.paths["logo"] and self.paths["participant"] and self.paths["presentation"] and self.paths["standing"]):
            print("One or more settings are missing")
            return
        pres = Presentation(self.paths["presentation"])
        create_slides_adjudicators(
            pres,
            self.dict_layouts,
            self.df_standings,
            self.df_participants,
            self.dict_logos,
            self.txt_title,
            self.display,
            self.danger_prevention
        )
        file_path = filedialog.asksaveasfilename(filetypes=[("PowerPoint file", "*.pptx")], title="Save as", defaultextension=".txt")
        if file_path:
            pres.save(file_path)
            print("Created file.")