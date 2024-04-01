import pandas as pd
import numpy as np
import functools
from pathlib import Path
from enum import Enum

class TeamMetric(Enum):
    AVERAGE_INDIVIDUAL_SPEAKS = "AISS"
    AVERAGE_MARGIN = "AWM"
    AVERAGE_TOTAL_SPEAKS = "ATSS"
    DRAW_STRENGTH_SCORE = "DSS"
    DRAW_STRENGTH_WIN = "DS"
    NUM_FIRSTS = "1sts"
    NUM_SECONDS = "2nds"
    NUM_THIRDS = "3rds"
    NUM_PULLUP = "SPu"
    NUM_IRONS = "Irons"
    POINTS = "Pts"
    STDEV_SPEAKS = "SSD"
    MARGINS = "Marg"
    TOTAL_SPEAKS = "Spk"
    BALLOTS = "Ballots"
    WHO_BEAT_WHOM = "WBW1"
    WINS = "Wins"

    @classmethod
    def values(cls):
        return list([e.value for e in cls])

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

description_team_metric = {
    TeamMetric.AVERAGE_INDIVIDUAL_SPEAKS: "Average individual speaker score",
    TeamMetric.AVERAGE_MARGIN: "Average margin",
    TeamMetric.AVERAGE_TOTAL_SPEAKS: "Average total speaker score",
    TeamMetric.DRAW_STRENGTH_SCORE: "Draw strength by total speaker score",
    TeamMetric.DRAW_STRENGTH_WIN: "Draw strength by wins",
    TeamMetric.NUM_FIRSTS: "Number of firsts",
    TeamMetric.NUM_SECONDS: "Number of seconds",
    TeamMetric.NUM_THIRDS: "Number of thirds",
    TeamMetric.NUM_PULLUP: "Number of times in pullup debates",
    TeamMetric.NUM_IRONS: "Number of times ironed",
    TeamMetric.POINTS: "Points",
    TeamMetric.STDEV_SPEAKS: "Speaker score standard deviation",
    TeamMetric.MARGINS: "Sum of margins",
    TeamMetric.TOTAL_SPEAKS: "Total speaker score",
    TeamMetric.BALLOTS: "Votes/ballots carried",
    TeamMetric.WHO_BEAT_WHOM: "Who-beat-whom",
    TeamMetric.WINS: "Wins",
}



def load_participant_debater(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="shift-jis")
    assert(set(["name", "category", "team", "institution"]).issubset(df.columns))
    return df

def load_participant_adjudicator(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    assert(set(["name", "institution"]).issubset(df.columns))
    return df

def load_logos(path: str) -> dict:
    dict_file = {}
    df = pd.read_csv(path)
    assert(all())
    for _, row in df.iterrows():
        dict_file[row["institution"]] = (Path(path).parent / row["path"]).resolve()
    assert(all(path.exists() for path in dict_file.values()))
    return dict_file

def load_standings_team(path: str, col_rank: str, metrics_show: list, metrics_hide: list):
    df = pd.read_csv(path)
    columns_necessary = ["team", col_rank, *metrics_show, *metrics_hide]
    assert(all(item in df.columns for item in columns_necessary))
    df_filter = df[columns_necessary].rename(columns={col_rank: "rank"}).dropna(subset=["rank"])
    # Erase unnecessary metrics
    for index, row in df_filter.iterrows():
        is_nan = False
        for i in range(len(metrics_hide)):
            if is_nan:
                df_filter.loc[index, metrics_hide[i]] = np.nan
            else:
                metrics_check = [*metrics_show, *metrics_hide[:i]]
                condition = functools.reduce(lambda x, y: x & y, [df_filter[metric] == row[metric] for metric in metrics_check])
                if len(df_filter[condition]) < 2:
                    df_filter.loc[index, metrics_hide[i]] = np.nan
                    is_nan = True
    return df_filter

def load_standings_debater(path: str, col_rank: str, metrics_show: list, metrics_hide: list):
    df = pd.read_csv(path)
    columns_necessary = ["name", col_rank, *metrics_show, *metrics_hide]
    assert(all(item in df.columns for item in columns_necessary))
    df_filter = df[columns_necessary].rename(columns={col_rank: "rank"}).dropna(subset=["rank"])
    # Erase unnecessary metrics
    for index, row in df_filter.iterrows():
        is_nan = False
        for i in range(len(metrics_hide)):
            if is_nan:
                df_filter.loc[index, metrics_hide[i]] = np.nan
            else:
                metrics_check = [*metrics_show, *metrics_hide[:i]]
                condition = functools.reduce(lambda x, y: x & y, [df_filter[metric] == row[metric] for metric in metrics_check])
                if len(df_filter[condition]) < 2:
                    df_filter.loc[index, metrics_hide[i]] = np.nan
                    is_nan = True
    return df_filter

def get_institution_team(df: pd.DataFrame, team: str):
    institutions = df[df["team"] == team]["institution"].dropna().unique()
    return institutions

def get_institution_debater(df: pd.DataFrame, name: str):
    debater = df[df["name"] == name]["institution"].dropna()
    return debater

def get_institution_adjudicator(df: pd.DataFrame, name: str):
    adj = df[df["name"] == name]["institution"].dropna()
    return adj

def get_institution_logo(dict_path: dict, institutions: list):
    try:
        paths = [str(dict_path[institution]) for institution in institutions]
    except Exception as e:
        raise Exception("failed to retreive institution")
    return paths