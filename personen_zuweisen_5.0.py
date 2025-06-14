"""
--------------------------------
Zuteilungs-Logik für Excel-Einsatzplan (BiPRO-Spezifikation)

Lesbare Regeln (Kurzfassung)
============================
* **Input**: Excel-Datei mit Reiter *Aufgaben* (A=Datum, B=Start, C=Ende, G=Kategorie, H=Personenzahl)
  und Reiter *Personen* (A=Name, B=Einsatzmöglichkeiten «;»-getrennt).
* **Output**:
  * **Einsatzplan_Zuteilung.csv** - Originalzeilen + Spalten >Eingeteilte Personen< & >Hinweis<
  * **Einsatzzeit_Statistik.csv** - Gesamt-Einsatzminuten pro Person
  Deutsches Format: Semikolon als Separator, Komma als Dezimaltrennzeichen.
* **Logik**: exakt wie vom Nutzer gefordert (Pausen, Backup, Sonderkategorien …). 

Aufruf
------
```bash
pip install pandas openpyxl
python einsatzplan_scheduler.py 2025-05-26_Einsatzplan.xlsx
```
"""

import sys
import pandas as pd
import datetime as dt
from typing import List, Dict, Tuple, Optional, Set
from dataclasses import dataclass, field

# ---------- Konfiguration ----------
SPECIAL_CATEGORIES: Set[str] = {"Regie", "Springer", "Orga-Springer"}
BREAK_PERSONS: Set[str] = {"Jan", "Ines"}
BREAK_MINUTES: int = 45
LUNCH_WINDOWS = [
    (dt.datetime(2025, 6, 4, 12, 30), dt.datetime(2025, 6, 4, 14, 0)),
    (dt.datetime(2025, 6, 5, 12, 0), dt.datetime(2025, 6, 5, 13, 30)),
]
LUNCH_FREE_MIN = 30
CSV_KWARGS = {"sep": ";", "decimal": ",", "encoding": "utf-8", "index": False}

# ---------- Hilfsfunktionen ----------

def parse_time(date: dt.date, time_value) -> dt.datetime:
    """Erzeugt datetime aus Datum + Excel-Zeit (string, float oder datetime.time)"""
    if pd.isna(time_value):
        raise ValueError("Zeitfeld fehlt")
    if isinstance(time_value, dt.datetime):
        t = time_value.time()
    elif isinstance(time_value, dt.time):
        t = time_value
    elif isinstance(time_value, (float, int)):
        total_seconds = round(24 * 60 * 60 * float(time_value))
        t = (dt.datetime.min + dt.timedelta(seconds=total_seconds)).time()
    else:
        time_str = str(time_value).strip()
        if ":" not in time_str:
            # Excel kann z.B. 1500 liefern → 15:00
            time_str = time_str.zfill(4)
            time_str = f"{time_str[:-2]}:{time_str[-2:]}"
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                t = dt.datetime.strptime(time_str, fmt).time()
                break
            except ValueError:
                continue
        else:
            raise ValueError(f"Unbekanntes Zeitformat: {time_value}")
    return dt.datetime.combine(date, t)

class Person:
    def __init__(self, name: str, categories: List[str]):
        self.name = name
        self.categories = categories
        # start, end, category, role ("main"|"backup")
        self.assignments: List[Tuple[dt.datetime, dt.datetime, str, str]] = []
        self.main_minutes = 0
        self.backup_minutes = 0

    @property
    def total_minutes(self) -> int:
        """Gesamte Einsatzzeit (nur Haupt-Minuten)."""
        return self.main_minutes

    def _overlap_minutes(self, s1, e1, s2, e2):
        latest_start = max(s1, s2)
        earliest_end = min(e1, e2)
        return max(0, int((earliest_end - latest_start).total_seconds() // 60))

    def _maintains_lunch(self, new_start: dt.datetime, new_end: dt.datetime) -> bool:
        for win_start, win_end in LUNCH_WINDOWS:
            if win_start.date() != new_start.date():
                continue
            busy = sum(
                self._overlap_minutes(s, e, win_start, win_end)
                for s, e, _, _ in self.assignments
            )
            busy += self._overlap_minutes(new_start, new_end, win_start, win_end)
            win_minutes = int((win_end - win_start).total_seconds() // 60)
            if busy > win_minutes - LUNCH_FREE_MIN:
                return False
        return True

    def is_available(self, start: dt.datetime, end: dt.datetime, category: str) -> bool:
        # Überschneidungen
		
        # Sonderregel: Frau Schrills am 05.06.2025 nicht einplanen
        if self.name.strip().lower().startswith('sandra') and start.date() == dt.date(2025, 6, 5):
            return False
			
        for a_start, a_end, a_cat, _ in self.assignments:
            if start < a_end and end > a_start:  # Überlapp
                if category in SPECIAL_CATEGORIES or a_cat in SPECIAL_CATEGORIES:
                    continue  # parallele Sonderkategorie erlaubt
                return False
        # Pause Jan/Ines
        if self.name.split()[0] in BREAK_PERSONS:
            for a_start, a_end, _, _ in self.assignments:
                if 0 <= (start - a_end).total_seconds() / 60 < BREAK_MINUTES:
                    return False
        # Mittagspause
        if not self._maintains_lunch(start, end):
            return False
        return True

    def assign(self, start: dt.datetime, end: dt.datetime, category: str, role: str = "main"):
        self.assignments.append((start, end, category, role))
        minutes = int((end - start).total_seconds() // 60)
        if role == "main":
            self.main_minutes += minutes
        else:
            self.backup_minutes += minutes


@dataclass
class Task:
    """Container für eine Aufgabe samt zugeteilter Personen"""
    index: int
    start: dt.datetime
    end: dt.datetime
    category: str
    required: int
    assigned: List[Person] = field(default_factory=list)
    backup: Optional[Person] = None


def balance_assignments(tasks: List[Task], people: Dict[str, Person]):
    """Schiebt Einsätze von stark zu wenig belasteten Personen."""
    while True:
        sorted_people = sorted(people.values(), key=lambda p: p.total_minutes)
        least = sorted_people[0]
        most = sorted_people[-1]
        if most.total_minutes - least.total_minutes <= 0:
            break

        moved = False
        for t in tasks:
            if most in t.assigned:
                if least.is_available(t.start, t.end, t.category) and t.category in least.categories:
                    t.assigned[t.assigned.index(most)] = least
                    for entry in most.assignments:
                        if entry[0] == t.start and entry[1] == t.end and entry[2] == t.category and entry[3] == "main":
                            most.assignments.remove(entry)
                            most.main_minutes -= int((t.end - t.start).total_seconds() // 60)
                            break
                    least.assign(t.start, t.end, t.category, role="main")
                    moved = True
                    break
        if not moved:
            break

def assign_backups(tasks: List[Task], people: Dict[str, Person]):
    """Weist Backups nur zu, wenn genügend Personen verfügbar sind."""
    for t in tasks:
        if len(t.assigned) < t.required:
            continue  # erst Hauptbesetzung sichern
        cand = select_candidates(
            list(people.values()), t.category, t.start, t.end, role="backup"
        )
        cand = [p for p in cand if p not in t.assigned]
        if cand:
            backup = cand[0]
            backup.assign(t.start, t.end, t.category, role="backup")
            t.backup = backup

def assign_main_round(tasks: List[Task], people: Dict[str, Person]):
    """Füllt offene Hauptrollen soweit moeglich."""
    for t in tasks:
        missing = t.required - len(t.assigned)
        if missing <= 0:
            continue
        cand = select_candidates(list(people.values()), t.category, t.start, t.end, role="main")
        cand = [p for p in cand if p not in t.assigned]
        selected = cand[:missing]
        for p in selected:
            p.assign(t.start, t.end, t.category, role="main")
            t.assigned.append(p)

# ---------- Kernlogik ----------

def select_candidates(
    pool: List[Person],
    category: str,
    start: dt.datetime,
    end: dt.datetime,
    role: str = "main",
) -> List[Person]:
    """Gibt verfügbare Personen sortiert nach Einsatzzeit zurück."""
    elig = [p for p in pool if category in p.categories and p.is_available(start, end, category)]
    if role == "main":
        elig.sort(key=lambda p: (p.total_minutes, p.main_minutes, p.name))
    else:
        elig.sort(key=lambda p: (p.total_minutes, p.backup_minutes, p.name))
    return elig

def main(xlsx_path: str):
    # Aufgaben laden - unabhängig von Kopfzeilentext
    tasks_df = pd.read_excel(xlsx_path, sheet_name="Aufgaben", header=0)
    persons_df = pd.read_excel(xlsx_path, sheet_name="Personen", header=0)

    # Personenliste aufbereiten
    people: Dict[str, Person] = {}
    for _, row in persons_df.iterrows():
        name = str(row.iloc[0]).strip()
        cats_raw = str(row.iloc[1]).replace(",", ";")
        cats = [c.strip() for c in cats_raw.split(";") if c.strip()]
        if cats:
            people[name] = Person(name, cats)

    # Ergebnis-Container
    out_persons: List[Optional[str]] = [None] * len(tasks_df)
    out_hints: List[str] = [""] * len(tasks_df)
    task_list: List[Task] = []

    for idx, row in tasks_df.iterrows():
        # Spalten über Index (robust gegen Kopfzeilen-Tipfehler)
        date_val = row.iat[0]
        start_val = row.iat[1]
        end_val = row.iat[2]
        category_val = row.iat[6] if len(row) > 6 else None
        req_val = row.iat[7] if len(row) > 7 else 0

        # Kategorie prüfen
        category = "" if pd.isna(category_val) else str(category_val).strip()
        if not category or category.lower() == "nan":
            out_persons[idx] = ""
            out_hints[idx] = "Kategorie leer - keine Einteilung"
            continue

        # Zeiten parsen
        try:
            date_obj = pd.to_datetime(date_val).date()
            start_dt = parse_time(date_obj, start_val)
            end_dt = parse_time(date_obj, end_val)
        except Exception as e:
            out_persons[idx] = ""
            out_hints[idx] = f"Zeitfehler: {e}"
            continue

        try:
            required = int(float(req_val)) if not pd.isna(req_val) else 0
        except ValueError:
            required = 0
        if required <= 0:
            out_persons[idx] = ""
            out_hints[idx] = "Personenanzahl 0 - keine Einteilung"
            continue

        task_list.append(Task(index=idx, start=start_dt, end=end_dt, category=category, required=required, assigned=[]))

    # vier Runden: drei ohne Backup, letzte mit Backup
    assign_main_round(task_list, people)
    balance_assignments(task_list, people)

    for _ in range(2):
        assign_main_round(task_list, people)
        balance_assignments(task_list, people)

    assign_main_round(task_list, people)
    balance_assignments(task_list, people)
    assign_backups(task_list, people)

    for task in task_list:
        names = [p.name for p in task.assigned]
        if task.backup:
            names.append(f"{task.backup.name} (Backup)")
        out_persons[task.index] = ", ".join(names)

        hint_parts = []
        if len(task.assigned) < task.required:
            hint_parts.append(
                f"Nur {len(task.assigned)} von {task.required} Personen gefunden"
            )
        if not task.backup:
            hint_parts.append("Kein Backup gefunden")
        out_hints[task.index] = "; ".join(hint_parts)

    tasks_df["Eingeteilte Personen"] = out_persons
    tasks_df["Hinweis"] = out_hints
    tasks_df.to_csv("Einsatzplan_Zuteilung.csv", **CSV_KWARGS)

    stats = pd.DataFrame([
        {"Person": p.name, "Einsatzzeit_Minuten": p.total_minutes} for p in people.values()
    ]).sort_values("Einsatzzeit_Minuten")
    stats.to_csv("Einsatzzeit_Statistik.csv", **CSV_KWARGS)

    print("CSV-Dateien erstellt: Einsatzplan_Zuteilung.csv, Einsatzzeit_Statistik.csv")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Aufruf: python einsatzplan_scheduler.py <Excel-Datei>")
        sys.exit(1)
    main(sys.argv[1])
