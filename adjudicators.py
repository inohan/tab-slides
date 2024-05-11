import pandas as pd
from pptx import Presentation
from out_ppt import create_slides_adjudicators, BlankSlideSettings
from pathlib import Path
import inquirer
from tkinter import filedialog
from collections import defaultdict
from exception import MissingColumnError, LogoError, enc
import logging
import json

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
        self.col_rank = None
        try:
            self.paths["logo"] = "./institution_logo.csv"
            self.load_logos()
        except Exception as e:
            print(e)

    def main_choices(self):
        txt_participant = f"Participant: {self.paths['participant']}" if self.paths["participant"] else "Participant: NOT LOADED"
        txt_logo = f"Logo: {self.paths['logo']}" if self.paths["logo"] else "Logo: NOT LOADED"
        txt_standing = f"Standing: {self.paths['standing']}" if self.paths["standing"] else "Standing: NOT LOADED"
        txt_presentation = f"Presentation: {self.paths['presentation']}" if self.paths["presentation"] else "Presentation: NOT LOADED"
        txt_title = f"Title: \"{self.txt_title.format('(nth)')}\""
        txt_display = f"Point display: \"{self.display}\""
        return [
            (txt_participant, "participant"),
            (txt_logo, "logo"),
            (txt_standing, "standing"),
            (txt_presentation, "presentation"),
            (txt_title, "title"),
            (txt_display, "display"),
            (f"Danger prevention: {self.danger_prevention.value}", "danger"),
            ("Create slide", "create"),
            ("Export settings", "export"),
            ("Quit", "quit")
        ]

    def main_menu(self):
        
        while (selected := inquirer.list_input("Configure adjudicator", choices=self.main_choices())) != "quit":
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
            elif selected == "export":
                self.save_settings()
    
    def prompt_path_logo(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select institution-logo file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.paths["logo"] = file_path
                self.load_logos()
            except Exception as e:
                print(e)
        else:
            print("Error: file does not exist.")

    def prompt_path_participant(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select adjudicator participant file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.paths["participant"] = file_path
                self.load_participants()
            except Exception as e:
                print(e)
        else:
            print("Error: file does not exist.")

    def prompt_path_standing(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select adjudicator standing file")
        if len(file_path) and Path(file_path).exists():
            try:
                df_load = pd.read_csv(file_path, encoding = enc)
                missing_columns = [item for item in ["name", "score"] if item not in df_load]
                if len(missing_columns) != 0:
                    raise MissingColumnError(f"Missing column(s) for {file_path}: {', '.join(missing_columns)}")
                # Selecting ranks
                self.col_rank = inquirer.list_input("Which column holds the rank?", choices = [c for c in df_load.columns if c not in ["name", "score"]])
                self.paths["standing"] = file_path
                self.load_standings()
            except Exception as e:
                logging.exception("What?")
        else:
            print("Error: file does not exist.")
    
    def prompt_path_presentation(self):
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint file", "*.pptx")], title="Select PowerPoint")
        if len(file_path) and Path(file_path).exists():
            try:
                pres = Presentation(file_path)
                choices = [(f"[{index}] {layout.name}", index) for index, layout in enumerate(pres.slide_layouts)]
                for i in range(2):
                    self.dict_layouts[i] = inquirer.list_input(f"Select slide for {i} institutions", choices=choices)
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

    def load_logos(self):
        df_logo = pd.read_csv(self.paths['logo'], encoding = enc)
        # Check for missing columns
        missing_columns = [item for item in ["institution", "path"] if item not in df_logo.columns]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {self.paths['logo']}: {', '.join(missing_columns)}")
        # Update dictionary
        for index, row in df_logo.iterrows():
            self.dict_logos[row["institution"]] = str(Path(row["path"])) if Path(row["path"]).exists() else None

    def load_participants(self):
        self.df_participants = pd.read_csv(self.paths['participant'], encoding = enc)
        # Check for missing columns
        missing_columns = [item for item in ["name", "institution"] if item not in self.df_participants.columns]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {self.paths['participant']}: {', '.join(missing_columns)}")

    def load_standings(self):
        df_load = pd.read_csv(self.paths["standing"], encoding = enc)
        # Selecting show metrics
        df_filter = df_load[["name", "score", self.col_rank]].rename(columns={self.col_rank: "rank"}).dropna(subset=["rank"]).sort_values(by=["rank", "name"], ascending=[False, True])
        self.df_standings = df_filter

    def save_settings(self):
        ret = {
            "type": "adjudicator",
            "path_participant": self.paths["participant"],
            "path_logo": self.paths["logo"],
            "standing": {
                "path": self.paths["standing"],
                "col_rank": self.col_rank
            },
            "presentation": {
                "path": self.paths["presentation"],
                "layouts": {
                    0: self.dict_layouts[0],
                    1: self.dict_layouts[1]
                }
            },
            "display": self.display,
            "title": self.txt_title,
            "danger_prevention": self.danger_prevention.value
        }
        with open("dump.json", mode="wt", encoding="utf-8") as f:
            json.dump(ret, f, ensure_ascii=False, indent=2)

    def load_settings(self, obj):
        self.paths = {
            "logo": obj["path_logo"],
            "participant": obj["path_participant"],
            "standing": obj["standing"]["path"],
            "presentation": obj["presentation"]["path"]
        }
        self.dict_layouts = {int(k): v for k, v in obj["presentation"]["layouts"].items()}
        self.display = obj["display"]
        self.col_rank = obj["standing"]["col_rank"]
        self.txt_title = obj["title"]
        self.danger_prevention = BlankSlideSettings(obj["danger_prevention"])
        self.load_logos()
        self.load_participants()
        self.load_standings()
        self.create_slides(obj["path_out"])

    def create_slides(self, file_path):
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
        if not file_path:
            file_path = filedialog.asksaveasfilename(filetypes=[("PowerPoint file", "*.pptx")], title="Save as", defaultextension=".txt")
        if file_path:
            pres.save(file_path)
            print("Created file.")