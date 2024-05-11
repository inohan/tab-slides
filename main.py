import os
import tkinter as tk
import teams
import speakers
import adjudicators
import inquirer
import json
from tkinter import filedialog

def load_all():
    path_json = filedialog.askopenfilename(filetypes=[("json file", "*.json")], title="Select json configs")
    is_first_query = True
    with open(path_json, mode="rt", encoding="utf-8") as f:
        data = json.load(f)		# JSONのファイル内容をdictに変換する。
    path_out = filedialog.asksaveasfilename(filetypes=[("PowerPoint file", "*.pptx")], title="Save as", defaultextension=".txt")
    for query in data:
        query["path_out"] = path_out
        if not is_first_query:
            query["presentation"]["path"] = path_out
        else:
            is_first_query = False
        if query["type"] == "team":
            teams.TeamSlidesBuilder().load_settings(query)
        elif query["type"] == "speaker":
            speakers.SpeakerSlidesBuilder().load_settings(query)
        elif query["type"] == "adjudicator":
            adjudicators.AdjudicatorSlidesBuilder().load_settings(query)
    print("All complete")


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    root = tk.Tk()
    root.withdraw()
    team_builder =teams.TeamSlidesBuilder()
    speaker_builder = speakers.SpeakerSlidesBuilder()
    adjudicator_builder = adjudicators.AdjudicatorSlidesBuilder()
    choices = [
        ("Team Slides", "team"),
        ("Speaker Slides", "speaker"),
        ("Adjudicator Slides", "adjudicator"),
        ("Bulk load from JSON", "json"),
        ("Quit", None)
    ]
    while selected := inquirer.list_input("Select which slide to create", choices=choices):
        if selected == "team":
            team_builder.main_menu()
        elif selected == "speaker":
            speaker_builder.main_menu()
        elif selected == "adjudicator":
            adjudicator_builder.main_menu()
        elif selected == "json":
            load_all()
    print("Done.")