import os
import json
import datetime
import requests
import openpyxl
from pathlib import Path

FOOTBALL_API_KEY = os.environ["FOOTBALL_API_KEY"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
COMPETITION_CODE = "WC"  # VM 2026

LAGNAMN = {
    "Mexico": "Mexiko",
    "South Africa": "Sydafrika",
    "USA": "USA",
    "Germany": "Tyskland",
    "France": "Frankrike",
    "Spain": "Spanien",
    "Brazil": "Brasilien",
    "Argentina": "Argentina",
    "England": "England",
    "Portugal": "Portugal",
    "Netherlands": "Nederländerna",
    "Morocco": "Marocko",
    "Japan": "Japan",
    "Australia": "Australien",
    "Saudi Arabia": "Saudiarabien",
    "South Korea": "Sydkorea",
    "Cameroon": "Kamerun",
    "Nigeria": "Nigeria",
    "Ghana": "Ghana",
    "Senegal": "Senegal",
    "Ecuador": "Ecuador",
    "Uruguay": "Uruguay",
    "Colombia": "Colombia",
    "Chile": "Chile",
}

def översätt(namn):
    return LAGNAMN.get(namn, namn)

def get_yesterdays_matches():
    url = f"https://api.football-data.org/v4/competitions/{COMPETITION_CODE}/matches"
    headers = {"X-Auth-Token": FOOTBALL_API_KEY}
    params = {"status": "FINISHED"}
    
    r = requests.get(url, headers=headers, params=params)
    r.raise_for_status()
    
    matches = r.json().get("matches", [])
    results = []
    for m in matches:
    home = översätt(m["homeTeam"]["shortName"] or m["homeTeam"]["name"])
    away = översätt(m["awayTeam"]["shortName"] or m["awayTeam"]["name"])
        home_score = m["score"]["fullTime"]["home"]
        away_score = m["score"]["fullTime"]["away"]
        if home_score > away_score:
            outcome = "1"
        elif home_score < away_score:
            outcome = "2"
        else:
            outcome = "X"
        results.append({
            "match": f"{home}–{away}",
            "home_score": home_score,
            "away_score": away_score,
            "outcome": outcome
        })
    return results

def load_tips():
    wb = openpyxl.load_workbook("VM_tips_2026.xlsx", read_only=True, data_only=True)
    ws = wb["Tips"]
    rows = list(ws.iter_rows(values_only=True))
    
    # Deltagarnamnen finns på rad index 3
    header_row = rows[3]
    participants = []
    for i, val in enumerate(header_row):
        if isinstance(val, str) and val.strip() and i >= 8:
            if val.strip() not in ("Rätt rad", "Torsdag 11 juni"):
                participants.append({"name": val.strip(), "col_index": i})
    
    # Matchrader börjar på index 4
    # Hoppa över rader som är datumrubriker (kolumn 0 är None eller inte ett tal)
    match_tips = []
    for row in rows[4:]:
        match_num = row[0]
        match_name = row[2]
        if not isinstance(match_num, (int, float)):
            continue  # datumrad eller tom rad
        if not match_name or not str(match_name).strip():
            continue
        tips_per_participant = {}
        for p in participants:
            tip = row[p["col_index"]]
            if tip in ("1", "X", "2", 1, 2):
                tips_per_participant[p["name"]] = str(int(tip)) if isinstance(tip, float) else str(tip)
        match_tips.append({
            "match": str(match_name).strip(),
            "tips": tips_per_participant
        })
    
    wb.close()
    return participants, match_tips
def calculate_scores(results, match_tips):
    scores = {}
    match_results = {}
    for result in results:
        for mt in match_tips:
            if any(team in mt["match"] for team in result["match"].split("–")):
                match_results[mt["match"]] = result["outcome"]
                for name, tip in mt["tips"].items():
                    if name not in scores:
                        scores[name] = {"correct": 0, "matches": []}
                    correct = tip == result["outcome"]
                    scores[name]["correct"] += int(correct)
                    scores[name]["matches"].append({
                        "match": mt["match"],
                        "tip": tip,
                        "outcome": result["outcome"],
                        "correct": correct
                    })
    return scores, match_results

def read_persona():
    try:
        return Path("persona.md").read_text(encoding="utf-8")
    except:
        return "Du är en entusiastisk tipsbloggare som älskar fotboll."

def generate_blog_post(results, scores, persona):
    if not results:
        return None
    yesterday = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d %B %Y")
    results_text = "\n".join([
        f"- {r['match']}: {r['home_score']}–{r['away_score']} (utfall: {r['outcome']})"
        for r in results
    ])
    scores_text = "\n".join([
        f"- {name}: {data['correct']} rätt av {len(data['matches'])}"
        for name, data in sorted(scores.items(), key=lambda x: -x[1]["correct"])
    ])
    johan_data = scores.get("J Nilsson", {})
    johan_detail = ""
    if johan_data:
        johan_detail = "Johans tips match för match:\n" + "\n".join([
            f"  {m['match']}: tippade {m['tip']}, utfall {m['outcome']} → {'✓' if m['correct'] else '✗'}"
            for m in johan_data.get("matches", [])
        ])
    prompt = f"""Skriv ett blogginlägg om gårdagens VM-tipsresultat ({yesterday}).

Matchresultat:
{results_text}

Poängställning efter gårdagens matcher:
{scores_text}

{johan_detail}

Blogginlägget ska:
- Vara på svenska
- Ha en catchy rubrik med datumet
- Kommentera matcherna med passion
- Lyfta fram tipsställningen med glimten i ögat
- Vara ca 300-400 ord
- Formateras som Markdown"""
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    body = {
        "model": "claude-sonnet-4-6",
        "max_tokens": 1000,
        "system": persona,
        "messages": [{"role": "user", "content": prompt}]
    }
    r = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=body)
    r.raise_for_status()
    return r.json()["content"][0]["text"]

def save_post(content):
    today = datetime.date.today().strftime("%Y-%m-%d")
    posts_dir = Path("_posts")
    posts_dir.mkdir(exist_ok=True)
    filename = posts_dir / f"{today}-tips-rapport.md"
    front_matter = f"""---
layout: post
title: "VM-tips {today}"
date: {today}
---

"""
    filename.write_text(front_matter + content, encoding="utf-8")
    print(f"Inlägg sparat: {filename}")

def main():
    print("Hämtar gårdagens matcher...")
    results = get_yesterdays_matches()
    if not results:
        print("Inga färdigspelade matcher igår – inget inlägg genereras.")
        return
    print(f"Hittade {len(results)} matcher.")
    participants, match_tips = load_tips()
    scores, match_results = calculate_scores(results, match_tips)
    persona = read_persona()
    print("Genererar blogginlägg...")
    post = generate_blog_post(results, scores, persona)
    if post:
        save_post(post)
        print("Klart!")

if __name__ == "__main__":
    main()
