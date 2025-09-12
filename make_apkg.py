import pandas as pd
import zipfile
import os
import sqlite3
import json
import shutil
import tempfile
import time
import random
import hashlib
from pathlib import Path
import sys

# -------- SETTINGS --------
CSV_PATH = "TT_PLU_codes.csv"
ZIP_PATH = "media.zip"
OUTPUT_DECK = "Produce_PLU_Name.apkg"
# --------------------------

def create_apkg(deck_name, cards, media_files, output_path):
    """Create an Anki .apkg file with cards = [(image_filename, answer)]."""
    tmpdir = tempfile.mkdtemp()
    col_path = Path(tmpdir) / "collection.anki2"

    # SQLite DB
    conn = sqlite3.connect(col_path)

    c = conn.cursor()

    # Create required tables
    c.execute("""CREATE TABLE revlog (
        id integer primary key,
        cid integer,
        usn integer,
        ease integer,
        ivl integer,
        lastIvl integer,
        factor integer,
        time integer,
        type integer
    )""")

    c.execute("""CREATE TABLE col (
        id integer primary key,
        crt integer, mod integer, scm integer, ver integer, dty integer, usn integer, ls integer,
        conf text, models text, decks text, dconf text, tags text)""")
    c.execute("""CREATE TABLE graves (
        usn integer,
        oid integer,
        type integer
    )""")

    c.execute("""CREATE TABLE notes (
        id integer primary key, guid text, mid integer, mod integer, usn integer, tags text,
        flds text, sfld integer, csum integer, flags integer, data text)""")

    c.execute("""CREATE TABLE cards (
        id integer primary key, nid integer, did integer, ord integer, mod integer, usn integer,
        type integer, queue integer, due integer, ivl integer, factor integer, reps integer,
        lapses integer, left integer, odue integer, odid integer, flags integer, data text)""")

    now = int(time.time() * 1000)
    deck_id = random.randrange(1 << 30, 1 << 31)
    model_id = random.randrange(1 << 30, 1 << 31)

    # Deck + model JSON
    decks = {
        str(deck_id): {
            "id": deck_id,
            "name": deck_name,
            "mod": now,
            "usn": 0,
            "desc": "",
            "dyn": 0,
            "collapsed": False,
            "extendNew": 0,
            "extendRev": 0,
            "conf": 1,
            "browserCollapsed": False,
            "newToday": [0, 0],
            "revToday": [0, 0],
            "lrnToday": [0, 0],
            "timeToday": [0, 0],
            "resched": True,
            "return": True
        }
    }

    models = {
        str(model_id): {
            "id": model_id,
            "name": "Basic (custom)",
            "type": 0,
            "mod": now,
            "usn": 0,
            "flds": [
                {"name": "Image", "ord": 0, "sticky": False, "rtl": False, "font": "Arial", "size": 20, "sortf": 0, "media": [], "description": ""},
                {"name": "Answer", "ord": 1, "sticky": False, "rtl": False, "font": "Arial", "size": 20, "sortf": 1, "media": [], "description": ""},
            ],
            "tmpls": [
                {"name": "Card 1", "ord": 0,
                 "qfmt": "<img src='{{Image}}'>",
                 "afmt": "{{FrontSide}}<hr id='answer'>{{Answer}}",
                 "did": None,
                 "bqfmt": "",
                 "bafmt": "",
                 "description": ""}
            ],
            "css": ".card { font-family: arial; font-size: 20px; text-align: center; color: black; background: white;}",
            "latexPre": "\\documentclass[12pt]{article}\\special{papersize=3in,5in}\\usepackage[utf8]{inputenc}\\usepackage{amssymb,amsmath}\\pagestyle{empty}\\setlength{\\parindent}{0in}\\begin{document}",
            "latexPost": "\\end{document}",
            "req": [[0, "all", [0]]],
            "tags": [],
            "vers": [],
            "type": 0,
            "did": None,
            "sortf": 0,
            "description": ""
        }
    }

    conf = {"nextPos": 1, "activeDecks": [deck_id], "curDeck": deck_id, "sortType": "noteCrt"}

    c.execute("INSERT INTO col VALUES (1,?,?,?,?,?,?,?, ?,?,?,?,?)",
              (now, now, now, 11, 0, 0, 0,
               json.dumps(conf), json.dumps(models), json.dumps(decks), "{}", "{}"))

    # Insert notes/cards
    for idx, (image, answer) in enumerate(cards):
        nid = now + idx
        guid = hashlib.md5(f"{nid}".encode()).hexdigest()
        flds = f"{image}\x1f{answer}"
        c.execute("INSERT INTO notes VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                  (nid, guid, model_id, now, 0, "", flds, 0, 0, 0, ""))
        cid = nid + 1
        c.execute("INSERT INTO cards VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                  (cid, nid, deck_id, 0, now, 0, 0, 0, idx + 1, 0, 0, 0, 0, 0, 0, 0, 0, ""))

    conn.commit()
    conn.close()

    # Copy media
    media_map = {}
    for i, (image, _) in enumerate(cards):
        src = media_files[image]
        dst_name = str(i)
        shutil.copy(src, Path(tmpdir) / dst_name)
        media_map[dst_name] = image

    with open(Path(tmpdir) / "media", "w", encoding="utf-8") as f:
        json.dump(media_map, f)

    # Zip into .apkg
    shutil.make_archive(output_path.replace(".apkg", ""), 'zip', tmpdir)
    shutil.move(output_path.replace(".apkg", "") + ".zip", output_path)
    shutil.rmtree(tmpdir)


def main():
    # Accept CSV path and output path as command-line arguments
    if len(sys.argv) >= 3:
        csv_path = sys.argv[1]
        output_deck = sys.argv[2]
    else:
        csv_path = CSV_PATH
        output_deck = OUTPUT_DECK

    # Load CSV
    df = pd.read_csv(csv_path)
    df.columns = df.columns.str.strip()  # Remove leading/trailing spaces from column names

    # Use images directly from the new folder
    extract_path = "media/media"
    # Get output filename prefix (without extension)
    output_file = os.path.basename(output_deck)
    prefix = os.path.splitext(output_file)[0] + "_"
    # Only use images with the correct prefix
    def image_sort_key(filename):
        import re
        m = re.match(re.escape(prefix) + r"image(\d+)\.png", filename)
        return int(m.group(1)) if m else float('inf')

    image_files = sorted([
        f for f in os.listdir(extract_path)
        if os.path.isfile(os.path.join(extract_path, f)) and f.startswith(prefix)
    ], key=image_sort_key)
    media_dict = {f: os.path.join(extract_path, f) for f in image_files}

    # Build cards: image on front, PLU and name on back
    cards = []
    for i in range(len(df)):
        if i < len(image_files):
            img = image_files[i]
            name = df.iloc[i]["Description"]
            plu = str(df.iloc[i]["No"])
            back = f"{plu} - {name}"
            cards.append((img, back))

    create_apkg("Produce - PLU & Name", cards, media_dict, output_deck)
    print("Deck created:", output_deck)



if __name__ == "__main__":
    main()
