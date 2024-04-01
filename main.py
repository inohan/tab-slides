import os
import tkinter as tk
import teams
import speakers
import adjudicators
import inquirer

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
        ("Quit", None)
    ]
    while selected := inquirer.list_input("Select which slide to create", choices=choices):
        if selected == "team":
            team_builder.main_menu()
        elif selected == "speaker":
            speaker_builder.main_menu()
        elif selected == "adjudicator":
            adjudicator_builder.main_menu()
    print("Done.")