import pandas as pd
from pptx import Presentation
from out_ppt import create_slides_debaters, BlankSlideSettings
from pathlib import Path
import inquirer
import numpy as np
import functools
from tkinter import filedialog
from collections import defaultdict
import copy
from enum import Enum
from exception import MissingColumnError, LogoError, enc
import logging
import json

class SpeakerSlidesBuilder():
    def __init__(self):
        self.df_participants = None
        self.df_standings = None
        self.dict_logos = defaultdict(None)
        self.dict_layouts = defaultdict(int)
        self.dict_metrics_display = copy.deepcopy(dict_speaker_metrics_display)
        self.pres = None
        self.metrics_show = []
        self.metrics_hide = []
        self.col_rank = None
        self.paths = {
            "logo": None,
            "participant": None,
            "standing": None,
            "presentation": None
        }
        self.txt_title = "{} Best Speaker"
        self.danger_prevention = BlankSlideSettings.CHANGE_RANK
        try:
            self.paths["logo"] = "./institution_logo.csv"
            self.load_logos()
        except Exception as e:
            print(e)

    def main_choices(self):
        txt_participant = f"Participant: {self.paths['participant']}" if self.paths["participant"] else "Participant: NOT LOADED"
        txt_logo = f"Logo: {self.paths['logo']}" if self.paths["logo"] else "Logo: NOT LOADED"
        txt_standing = f"Standing: {self.paths['standing']} | {'->'.join([m.value for m in self.metrics_show])}" if self.paths["standing"] else "Standing: NOT LOADED"
        txt_presentation = f"Presentation: {self.paths['presentation']}" if self.paths["presentation"] else "Presentation: NOT LOADED"
        txt_title = f"Title: \"{self.txt_title.format('(nth)')}\""
        txt_display_metrics = f"Metrics display"
        if len(self.metrics_hide):
            txt_standing += f"(->{'->'.join([m.value for m in self.metrics_hide])})"
        return [
            (txt_participant, "participant"),
            (txt_logo, "logo"),
            (txt_standing, "standing"),
            (txt_presentation, "presentation"),
            (txt_title, "title"),
            (txt_display_metrics, "display"),
            (f"Danger prevention: {self.danger_prevention.value}", "danger"),
            ("Create slide", "create"),
            ("Export settings", "export"),
            ("Quit", "quit")
        ]

    def main_menu(self):
        while (selected := inquirer.list_input("Configure speaker", choices=self.main_choices())) != "quit":
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
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select team/debater participant file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.paths["participant"] = file_path
                self.load_participants()
            except Exception as e:
                print(e)
        else:
            print("Error: file does not exist.")

    def prompt_path_standing(self):
        file_path = filedialog.askopenfilename(filetypes=[("csv file", "*.csv")], title="Select debater standing file")
        if len(file_path) and Path(file_path).exists():
            try:
                self.metrics_show = []
                self.metrics_hide = []
                df_load = pd.read_csv(file_path, encoding = enc)
                if "name" not in df_load.columns:
                    raise MissingColumnError("Missing Column 'name'")
                columns_metric = [SpeakerMetric(e) for e in SpeakerMetric.values() if e in df_load.columns]
                # Selecting ranks
                self.col_rank = inquirer.list_input("Which column holds the rank?", choices = [c for c in df_load.columns if c not in [e.value for e in columns_metric]])
                # Selecting show metrics
                while len(metrics_remaining := list(set(columns_metric) - set(self.metrics_hide) - set(self.metrics_show))):
                    choices = [(f"{m.value} ({description_speaker_metric[SpeakerMetric(m)]})", SpeakerMetric(m)) for m in metrics_remaining] + [("(Next)", None)]
                    selected = inquirer.list_input(f"Select all metrics to always show: {' -> '.join([m.value for m in self.metrics_show])}", choices = choices)
                    if selected:
                        self.metrics_show.append(selected)
                    else:
                        break
                # Selecting hide metrics
                while len(metrics_remaining := list(set(columns_metric) - set(self.metrics_hide) - set(self.metrics_show))):
                    choices = [(f"{m.value} ({description_speaker_metric[SpeakerMetric(m)]})", SpeakerMetric(m)) for m in metrics_remaining] + [("(Done)", None)]
                    selected = inquirer.list_input(f"Select all metrics to show only when necessary: {' -> '.join([m.value for m in self.metrics_hide])}", choices = choices)
                    if selected:
                        self.metrics_hide.append(selected)
                    else:
                        break
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
        self.txt_title = inquirer.text("Title text (e.g. \"(nth) Best Speaker\")", default="(nth) Best Speaker").replace("(nth)", "{}")

    def prompt_metric_display(self):
        choices = [(f"{description_speaker_metric[key]} ({key.value}): \"{value}\"", key) for key, value in self.dict_metrics_display.items()] + [("Reset", "reset"), ("Back", None)]
        selected = inquirer.list_input("Which display to change?", choices=choices)
        if selected and selected != "reset":
            self.dict_metrics_display[selected] = inquirer.text(f"Enter format for {description_speaker_metric[selected]}")
            self.prompt_metric_display()
        elif selected and selected == "reset":
            self.dict_metrics_display = copy.deepcopy(dict_speaker_metrics_display)
        else:
            pass

    def prompt_danger_prevention(self):
        choices = [(e.value, e) for e in BlankSlideSettings]
        self.danger_prevention = inquirer.list_input("Select danger prevention settings", default=self.danger_prevention, choices = choices)

    def load_logos(self):
        df_logo = pd.read_csv(self.paths["logo"], encoding = enc)
        # Check for missing columns
        missing_columns = [item for item in ["institution", "path"] if item not in df_logo.columns]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {self.paths['logo']}: {', '.join(missing_columns)}")
        # Update dictionary
        for index, row in df_logo.iterrows():
            self.dict_logos[row["institution"]] = str(Path(row["path"])) if Path(row["path"]).exists() else None

    def load_participants(self):
        self.df_participants = pd.read_csv(self.paths["participant"], encoding = enc)
        # Check for missing columns
        missing_columns = [item for item in ["name", "institution"] if item not in self.df_participants.columns]
        if len(missing_columns) != 0:
            raise MissingColumnError(f"Missing column(s) for {self.paths['participant']}: {', '.join(missing_columns)}")

    def load_standings(self):
        df_load = pd.read_csv(self.paths["standing"], encoding = enc)
        columns_extract = ["name", self.col_rank, *[m.value for m in self.metrics_show], *[m.value for m in self.metrics_hide]]
        df_filter = df_load[columns_extract].rename(columns={self.col_rank: "rank"})
        # Remove unnecessary metrics
        for index, row in df_filter.iterrows():
            is_nan = False
            for i in range(len(self.metrics_hide)):
                if is_nan:
                    df_filter.loc[index, self.metrics_hide[i]] = np.nan
                else:
                    metrics_check = [*self.metrics_show, *self.metrics_hide[:i]]
                    condition = functools.reduce(lambda x, y: x & y, [df_filter[metric.value] == row[metric.value] for metric in metrics_check])
                    if len(df_filter[condition]) < 2:
                        df_filter.loc[index, self.metrics_hide[i]] = np.nan
                        is_nan = True
        self.df_standings = df_filter.dropna(subset=["rank"]).sort_values(by=["rank", "name"], ascending=[False, True])

    def save_settings(self):
        ret = {
            "type": "speaker",
            "path_participant": self.paths["participant"],
            "path_logo": self.paths["logo"],
            "standing": {
                "path": self.paths["standing"],
                "col_rank": self.col_rank,
                "metrics_show": [m.value for m in self.metrics_show],
                "metrics_hide": [m.value for m in self.metrics_hide]
            },
            "presentation": {
                "path": self.paths["presentation"],
                "layouts": {
                    0: self.dict_layouts[0],
                    1: self.dict_layouts[1]
                }
            },
            "title": self.txt_title,
            "danger_prevention": self.danger_prevention.value
        }
        disp_metrics = {}
        for k, v in self.dict_metrics_display.items():
            if v != dict_speaker_metrics_display[k]:
                disp_metrics[k.value] = v
        ret["metrics_display"] = disp_metrics
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
        self.dict_metrics_display = copy.deepcopy(dict_speaker_metrics_display)
        for k, v in obj["metrics_display"].items():
            self.dict_metrics_display[SpeakerMetric(k)] = v
        self.col_rank = obj["standing"]["col_rank"]
        self.metrics_show = [SpeakerMetric(m) for m in obj["standing"]["metrics_show"]]
        self.metrics_hide = [SpeakerMetric(m) for m in obj["standing"]["metrics_hide"]]
        self.txt_title = obj["title"]
        self.danger_prevention = BlankSlideSettings(obj["danger_prevention"])
        self.load_logos()
        self.load_participants()
        self.load_standings()
        self.create_slides(obj["path_out"])

    def create_slides(self, file_path = None):
        if not (self.paths["logo"] and self.paths["participant"] and self.paths["presentation"] and self.paths["standing"]):
            print("One or more settings are missing")
            return
        pres = Presentation(self.paths["presentation"])
        create_slides_debaters(
            pres,
            self.dict_layouts,
            self.df_standings,
            self.df_participants,
            self.dict_logos,
            self.txt_title,
            self.metrics_show,
            self.metrics_hide,
            self.dict_metrics_display,
            self.danger_prevention
        )
        if not file_path:
            file_path = filedialog.asksaveasfilename(filetypes=[("PowerPoint file", "*.pptx")], title="Save as", defaultextension=".txt")
        if file_path:
            pres.save(file_path)
            print("Created file.")



class SpeakerMetric(Enum):
    TEAM_POINTS = "Team"
    AVERAGE = "Avg"
    NUM_SPEECHES = "Num"
    SPEECH_RANKS = "SRank"
    STDEV = "Stdev"
    TOTAL = "Total"
    TRIM = "Trim"

    @classmethod
    def values(cls):
        return list([e.value for e in cls])

description_speaker_metric = {
    SpeakerMetric.TEAM_POINTS: "Team points",
    SpeakerMetric.AVERAGE: "Average",
    SpeakerMetric.NUM_SPEECHES: "Number of speeches given",
    SpeakerMetric.SPEECH_RANKS: "Speech ranks",
    SpeakerMetric.STDEV: "Standard deviation",
    SpeakerMetric.TOTAL: "Total",
    SpeakerMetric.TRIM: "Trimmed mean (high-low drop)"
}

dict_speaker_metrics_display = {
    SpeakerMetric.TEAM_POINTS: "{} points",
    SpeakerMetric.AVERAGE: "avg. {} spks",
    SpeakerMetric.NUM_SPEECHES: "{} speeches",
    SpeakerMetric.SPEECH_RANKS: "{} speech ranks",
    SpeakerMetric.STDEV: "stdev. {}",
    SpeakerMetric.TOTAL: "{} spks",
    SpeakerMetric.TRIM: "{} spks (trimmed)"
}