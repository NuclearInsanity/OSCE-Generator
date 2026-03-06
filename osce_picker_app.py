from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from html import escape
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


DEFAULT_WORKBOOK = Path(
    "/Users/alexanderpanayotouennes/Desktop/Codex/OSCE station - Stem and marking rubric/OSCE's accumulated 6.3.26 with stem and rubric.xlsx"
)
DEFAULT_OUTPUT = Path(__file__).with_name("AFRM OSCE Picker.html")
DEFAULT_SITE_OUTPUT = Path(__file__).with_name("docs").joinpath("index.html")
XML_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def normalize_text(value: str) -> str:
    return re.sub(r"[ \t]+", " ", value or "").strip()


def normalize_header(value: str) -> str:
    lowered = normalize_text(value).lower()
    lowered = lowered.replace("#", " number ")
    lowered = lowered.replace("-", " ")
    cleaned = re.sub(r"[^a-z0-9 ]+", "", lowered)
    return normalize_text(cleaned)


def split_subtopics(value: str) -> tuple[str, ...]:
    parts = []
    for item in re.split(r",|;", value or ""):
        cleaned = normalize_text(item)
        if cleaned and cleaned not in parts:
            parts.append(cleaned)
    return tuple(parts)


def column_ref_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    total = 0
    for char in letters:
        total = total * 26 + (ord(char.upper()) - 64)
    return max(total - 1, 0)


def load_shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    shared = []
    for item in root.findall("main:si", XML_NS):
        text_parts = []
        for node in item.iterfind(".//main:t", XML_NS):
            text_parts.append(node.text or "")
        shared.append("".join(text_parts))
    return shared


def worksheet_rows(path: Path) -> list[list[str]]:
    with ZipFile(path) as archive:
        shared_strings = load_shared_strings(archive)
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        rows: list[list[str]] = []
        for row in root.findall(".//main:sheetData/main:row", XML_NS):
            values_by_index: dict[int, str] = {}
            for cell in row.findall("main:c", XML_NS):
                index = column_ref_to_index(cell.attrib.get("r", "A1"))
                cell_type = cell.attrib.get("t")
                raw_value = cell.find("main:v", XML_NS)
                inline = cell.find("main:is", XML_NS)
                value = ""
                if cell_type == "s" and raw_value is not None and raw_value.text:
                    value = shared_strings[int(raw_value.text)]
                elif cell_type == "inlineStr" and inline is not None:
                    text_parts = []
                    for node in inline.iterfind(".//main:t", XML_NS):
                        text_parts.append(node.text or "")
                    value = "".join(text_parts)
                elif raw_value is not None and raw_value.text is not None:
                    value = raw_value.text
                values_by_index[index] = value
            if values_by_index:
                max_index = max(values_by_index)
                rows.append([values_by_index.get(i, "") for i in range(max_index + 1)])
        return rows


def detect_header_row(rows: list[list[str]]) -> int:
    for index, row in enumerate(rows[:10]):
        normalized = {normalize_header(cell) for cell in row if normalize_text(cell)}
        if "year" in normalized and "question" in normalized:
            return index
    raise ValueError("Could not find the workbook header row.")


@dataclass(frozen=True)
class Station:
    row_number: int
    year: str
    station_type: str
    question_number: str
    question: str
    question_polished: str
    examination: str
    main_topic: str
    subtopics_raw: str
    subtopics: tuple[str, ...]
    stem: str
    marking_rubric: str
    answer: str

    @property
    def title(self) -> str:
        parts = [self.year or "Unknown year"]
        if self.station_type:
            parts.append(self.station_type)
        if self.question_number:
            parts.append(f"Q{self.question_number}")
        return " | ".join(parts)

    def to_payload(self) -> dict[str, object]:
        payload = asdict(self)
        payload["title"] = self.title
        payload["stem"] = self.stem or self.question or "No stem text available."
        payload["original_question"] = self.question or "No original question text available."
        payload["display_question"] = self.question_polished or self.question or "No question text available."
        return payload


def load_stations(path: Path) -> list[Station]:
    rows = worksheet_rows(path)
    header_row = detect_header_row(rows)
    headers = [normalize_header(cell) for cell in rows[header_row]]

    def cell(row: list[str], header_name: str) -> str:
        try:
            idx = headers.index(header_name)
        except ValueError:
            return ""
        if idx >= len(row):
            return ""
        return (row[idx] or "").strip()

    def first_available(row: list[str], *header_names: str) -> str:
        for header_name in header_names:
            value = cell(row, header_name)
            if value:
                return value
        return ""

    stations: list[Station] = []
    for raw_row in rows[header_row + 1 :]:
        question = cell(raw_row, "question")
        year = cell(raw_row, "year")
        if not question and not year:
            continue
        subtopics_raw = cell(raw_row, "sub topic")
        stations.append(
            Station(
                row_number=len(stations) + 1,
                year=year,
                station_type=cell(raw_row, "type mockrecallofficial"),
                question_number=cell(raw_row, "question number"),
                question=question,
                question_polished=first_available(raw_row, "question polished", "questionpolished"),
                examination=cell(raw_row, "physical exam if present"),
                main_topic=cell(raw_row, "main topic"),
                subtopics_raw=subtopics_raw,
                subtopics=split_subtopics(subtopics_raw),
                stem=first_available(raw_row, "predicted stem", "stem"),
                marking_rubric=first_available(raw_row, "marking rubric", "markingrubric"),
                answer=cell(raw_row, "gpt answers"),
            )
        )
    if not stations:
        raise ValueError("No stations were found in the workbook.")
    return stations


def safe_json_for_html(value: object) -> str:
    return (
        json.dumps(value, ensure_ascii=False)
        .replace("</", "<\\/")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def html_template(stations: list[Station], workbook_path: Path) -> str:
    payload = [station.to_payload() for station in stations]
    data_json = safe_json_for_html(payload)
    workbook_name = escape(workbook_path.name)
    built_at = escape(datetime.now().strftime("%d %b %Y %H:%M"))
    station_count = str(len(stations))
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>AFRM OSCE Picker</title>
  <style>
    :root {{
      --bg: #edf3f0;
      --panel: #fbfcfb;
      --panel-alt: #f4f8f6;
      --text: #17352e;
      --muted: #59726a;
      --line: #d7e2dd;
      --accent: #1f6b58;
      --accent-dark: #154f41;
      --soft: #dce9e3;
      --shadow: 0 18px 40px rgba(19, 39, 34, 0.08);
    }}

    * {{
      box-sizing: border-box;
    }}

    body {{
      margin: 0;
      min-height: 100vh;
      font-family: "SF Pro Text", "Helvetica Neue", Helvetica, Arial, sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top right, rgba(31, 107, 88, 0.12), transparent 28rem),
        linear-gradient(180deg, #f3f7f5 0%, var(--bg) 100%);
    }}

    .shell {{
      max-width: 1480px;
      margin: 0 auto;
      padding: 28px;
    }}

    .hero {{
      display: flex;
      justify-content: space-between;
      gap: 24px;
      align-items: flex-start;
      margin-bottom: 18px;
    }}

    .hero h1 {{
      margin: 0;
      font-size: clamp(2rem, 3vw, 2.8rem);
      line-height: 1;
    }}

    .hero p {{
      margin: 10px 0 0;
      color: var(--muted);
      max-width: 54rem;
      line-height: 1.5;
    }}

    .workbook {{
      min-width: 20rem;
      padding: 16px 18px;
      background: rgba(255, 255, 255, 0.75);
      border: 1px solid rgba(215, 226, 221, 0.9);
      border-radius: 18px;
      box-shadow: var(--shadow);
    }}

    .workbook strong {{
      display: block;
      margin-bottom: 4px;
      font-size: 0.95rem;
    }}

    .workbook span {{
      display: block;
      color: var(--muted);
      font-size: 0.9rem;
      line-height: 1.45;
      word-break: break-word;
    }}

    .card {{
      background: rgba(251, 252, 251, 0.88);
      border: 1px solid rgba(215, 226, 221, 0.92);
      border-radius: 24px;
      box-shadow: var(--shadow);
      backdrop-filter: blur(8px);
    }}

    .filters {{
      padding: 22px;
      margin-bottom: 18px;
    }}

    .filter-grid {{
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 14px;
    }}

    .field {{
      display: flex;
      flex-direction: column;
      gap: 7px;
    }}

    .field label {{
      font-size: 0.84rem;
      font-weight: 700;
      letter-spacing: 0.02em;
      color: #21463e;
    }}

    select,
    input {{
      width: 100%;
      min-height: 46px;
      padding: 11px 13px;
      font: inherit;
      color: var(--text);
      background: white;
      border: 1px solid var(--line);
      border-radius: 14px;
      outline: none;
    }}

    select:focus,
    input:focus {{
      border-color: rgba(31, 107, 88, 0.55);
      box-shadow: 0 0 0 4px rgba(31, 107, 88, 0.12);
    }}

    .actions {{
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: center;
      margin-top: 18px;
      flex-wrap: wrap;
    }}

    .button-row {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }}

    button {{
      border: 0;
      border-radius: 999px;
      padding: 12px 16px;
      font: inherit;
      cursor: pointer;
      transition: transform 120ms ease, background 120ms ease, opacity 120ms ease;
    }}

    button:hover {{
      transform: translateY(-1px);
    }}

    .primary {{
      background: var(--accent);
      color: white;
      font-weight: 700;
    }}

    .primary:hover {{
      background: var(--accent-dark);
    }}

    .secondary {{
      background: var(--soft);
      color: var(--text);
    }}

    .ghost {{
      background: transparent;
      color: var(--accent);
      border: 1px solid rgba(31, 107, 88, 0.22);
    }}

    .pool {{
      color: var(--muted);
      font-size: 0.95rem;
    }}

    .layout {{
      display: grid;
      grid-template-columns: minmax(0, 1.7fr) minmax(360px, 1fr);
      gap: 18px;
      align-items: start;
    }}

    .stem-card,
    .side-card {{
      padding: 24px;
    }}

    .stem-card {{
      display: block;
    }}

    .side-stack {{
      display: grid;
      grid-template-rows: minmax(0, 0.9fr) minmax(0, 1.1fr);
      gap: 18px;
      min-height: 0;
    }}

    .section-head {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 10px;
    }}

    .section-head h2 {{
      margin: 0;
      font-size: 1.2rem;
    }}

    .meta {{
      color: var(--muted);
      font-size: 0.94rem;
      line-height: 1.6;
      margin-bottom: 18px;
      white-space: pre-line;
    }}

    .copy {{
      color: var(--text);
      background: var(--panel-alt);
      border: 1px solid rgba(215, 226, 221, 0.85);
      border-radius: 18px;
      padding: 16px 18px;
      line-height: 1.62;
      white-space: pre-wrap;
      overflow: auto;
      min-height: 0;
      flex: 1;
    }}

    .stem-primary {{
      margin-bottom: 14px;
    }}

    .stem-secondary {{
      flex: 0 0 auto;
    }}

    .stem-primary .copy {{
      min-height: 22rem;
      overflow: visible;
    }}

    .stem-secondary .copy {{
      min-height: 12rem;
      overflow: visible;
    }}

    .hint {{
      color: var(--muted);
      margin: 6px 0 12px;
      line-height: 1.5;
    }}

    .panel-block {{
      display: flex;
      flex-direction: column;
      min-height: 0;
    }}

    .subhead {{
      font-size: 0.85rem;
      font-weight: 700;
      letter-spacing: 0.02em;
      color: #33594f;
      margin: 0 0 8px;
    }}

    .rubric-block {{
      margin-bottom: 14px;
      min-height: 0;
    }}

    .rubric-block .copy {{
      min-height: 12rem;
    }}

    .answer-block {{
      flex: 1;
      min-height: 0;
    }}

    .answer-block .copy {{
      min-height: 16rem;
    }}

    @media (max-width: 1120px) {{
      .filter-grid {{
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }}

      .layout {{
        grid-template-columns: 1fr;
      }}

      .side-stack {{
        grid-template-rows: auto;
      }}
    }}

    @media (max-width: 720px) {{
      .shell {{
        padding: 18px;
      }}

      .hero {{
        flex-direction: column;
      }}

      .filter-grid {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <div>
        <h1>AFRM OSCE Picker</h1>
        <p>Filter by year, topic, sub-topic, or station focus, then draw a random station. The stem is shown first, questions stay hidden until you reveal them, and the marking rubric appears together with the answer.</p>
      </div>
      <div class="workbook">
        <strong>{workbook_name}</strong>
        <span>Built from the latest workbook export.</span>
        <span>{station_count} stations • Updated {built_at}</span>
      </div>
    </section>

    <section class="card filters">
      <div class="filter-grid">
        <div class="field">
          <label for="yearFilter">Year</label>
          <select id="yearFilter"></select>
        </div>
        <div class="field">
          <label for="typeFilter">Type</label>
          <select id="typeFilter"></select>
        </div>
        <div class="field">
          <label for="topicFilter">Main topic</label>
          <input id="topicFilter" list="topicOptions" placeholder="Any topic">
          <datalist id="topicOptions"></datalist>
        </div>
        <div class="field">
          <label for="subtopicFilter">Sub-topic</label>
          <input id="subtopicFilter" list="subtopicOptions" placeholder="Any sub-topic">
          <datalist id="subtopicOptions"></datalist>
        </div>
        <div class="field">
          <label for="examFilter">Specific examination</label>
          <input id="examFilter" list="examOptions" placeholder="Any examination">
          <datalist id="examOptions"></datalist>
        </div>
      </div>
      <div class="actions">
        <div class="button-row">
          <button class="primary" id="pickFilteredButton">Pick from filters</button>
          <button class="secondary" id="pickAllButton">Random from all</button>
          <button class="ghost" id="clearFiltersButton">Clear filters</button>
        </div>
        <div class="pool" id="poolLabel"></div>
      </div>
    </section>

    <section class="layout">
      <article class="card stem-card">
        <div class="section-head">
          <h2>Stem</h2>
        </div>
        <div class="meta" id="stationMeta"></div>
        <div class="stem-primary">
          <p class="subhead">Predicted stem</p>
          <div class="copy" id="stemBody">Choose a random station to begin.</div>
        </div>
        <div class="stem-secondary">
          <p class="subhead">Original source question</p>
          <div class="copy" id="originalQuestionBody">The original question will appear here.</div>
        </div>
      </article>

      <section class="side-stack">
        <article class="card side-card panel-block">
          <div class="section-head">
            <h2>Questions</h2>
            <button class="secondary" id="toggleQuestionsButton">Show questions</button>
          </div>
          <p class="hint" id="questionsHint">Questions hidden. Click to reveal the station brief.</p>
          <div class="copy" id="questionsBody">The question prompt is currently hidden.</div>
        </article>

        <article class="card side-card panel-block">
          <div class="section-head">
            <h2>Marking rubric and answer</h2>
            <button class="primary" id="toggleAnswerButton">Show answer</button>
          </div>
          <p class="hint" id="answerHint">Answer hidden. Click "Show answer" to reveal the marking rubric and GPT answer.</p>
          <div class="rubric-block">
            <p class="subhead">Marking rubric</p>
            <div class="copy" id="rubricBody">The marking rubric is currently hidden.</div>
          </div>
          <div class="answer-block">
            <p class="subhead">GPT answer</p>
            <div class="copy" id="answerBody">The answer is currently hidden.</div>
          </div>
        </article>
      </section>
    </section>
  </div>

  <script>
    const stations = {data_json};
    const state = {{
      currentStation: null,
      questionsVisible: false,
      answerVisible: false,
    }};

    const elements = {{
      yearFilter: document.getElementById("yearFilter"),
      typeFilter: document.getElementById("typeFilter"),
      topicFilter: document.getElementById("topicFilter"),
      subtopicFilter: document.getElementById("subtopicFilter"),
      examFilter: document.getElementById("examFilter"),
      topicOptions: document.getElementById("topicOptions"),
      subtopicOptions: document.getElementById("subtopicOptions"),
      examOptions: document.getElementById("examOptions"),
      poolLabel: document.getElementById("poolLabel"),
      stationMeta: document.getElementById("stationMeta"),
      stemBody: document.getElementById("stemBody"),
      originalQuestionBody: document.getElementById("originalQuestionBody"),
      questionsHint: document.getElementById("questionsHint"),
      questionsBody: document.getElementById("questionsBody"),
      answerHint: document.getElementById("answerHint"),
      rubricBody: document.getElementById("rubricBody"),
      answerBody: document.getElementById("answerBody"),
      toggleQuestionsButton: document.getElementById("toggleQuestionsButton"),
      toggleAnswerButton: document.getElementById("toggleAnswerButton"),
      pickFilteredButton: document.getElementById("pickFilteredButton"),
      pickAllButton: document.getElementById("pickAllButton"),
      clearFiltersButton: document.getElementById("clearFiltersButton"),
    }};

    function setSelectOptions(select, values) {{
      select.innerHTML = "";
      const anyOption = document.createElement("option");
      anyOption.value = "Any";
      anyOption.textContent = "Any";
      select.appendChild(anyOption);
      values.forEach((value) => {{
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
      }});
      select.value = "Any";
    }}

    function setDatalistOptions(list, values) {{
      list.innerHTML = "";
      values.forEach((value) => {{
        const option = document.createElement("option");
        option.value = value;
        list.appendChild(option);
      }});
    }}

    function uniqueSorted(values, numeric = false) {{
      const unique = [...new Set(values.filter((value) => value && value.trim()))];
      if (numeric) {{
        unique.sort((left, right) => {{
          const leftNumber = Number(left);
          const rightNumber = Number(right);
          if (!Number.isNaN(leftNumber) && !Number.isNaN(rightNumber)) {{
            return leftNumber - rightNumber;
          }}
          return left.localeCompare(right);
        }});
      }} else {{
        unique.sort((left, right) => left.localeCompare(right));
      }}
      return unique;
    }}

    function populateFilters() {{
      setSelectOptions(elements.yearFilter, uniqueSorted(stations.map((station) => station.year), true));
      setSelectOptions(elements.typeFilter, uniqueSorted(stations.map((station) => station.station_type)));
      setDatalistOptions(elements.topicOptions, uniqueSorted(stations.map((station) => station.main_topic)));
      setDatalistOptions(elements.subtopicOptions, uniqueSorted(stations.flatMap((station) => station.subtopics)));
      setDatalistOptions(elements.examOptions, uniqueSorted(stations.map((station) => station.examination)));
      elements.topicFilter.value = "";
      elements.subtopicFilter.value = "";
      elements.examFilter.value = "";
    }}

    function normalized(value) {{
      return (value || "").trim().toLowerCase();
    }}

    function filteredStations() {{
      const year = elements.yearFilter.value;
      const type = elements.typeFilter.value;
      const topic = normalized(elements.topicFilter.value);
      const subtopic = normalized(elements.subtopicFilter.value);
      const exam = normalized(elements.examFilter.value);

      return stations.filter((station) => {{
        if (year !== "Any" && station.year !== year) {{
          return false;
        }}
        if (type !== "Any" && station.station_type !== type) {{
          return false;
        }}
        if (topic && !normalized(station.main_topic).includes(topic)) {{
          return false;
        }}
        if (exam && !normalized(station.examination).includes(exam)) {{
          return false;
        }}
        if (subtopic) {{
          const match = station.subtopics.some((entry) => normalized(entry).includes(subtopic));
          if (!match) {{
            return false;
          }}
        }}
        return true;
      }});
    }}

    function updatePoolLabel() {{
      const matches = filteredStations().length;
      elements.poolLabel.textContent = `${{matches}} matching stations / ${{stations.length}} total`;
    }}

    function renderMeta(station) {{
      const subtopicsPreview = station.subtopics.slice(0, 6).join(", ");
      const subtopicsText = station.subtopics.length > 6 ? `${{subtopicsPreview}}, ...` : subtopicsPreview;
      elements.stationMeta.textContent = [
        station.title,
        `Main topic: ${{station.main_topic || "Not specified"}}`,
        `Specific examination: ${{station.examination || "Not specified"}}`,
        `Sub-topics: ${{subtopicsText || "Not specified"}}`,
      ].join("\\n");
    }}

    function renderPanels() {{
      const station = state.currentStation;
      if (!station) {{
        elements.stationMeta.textContent = "";
        elements.stemBody.textContent = "Choose a random station to begin.";
        elements.originalQuestionBody.textContent = "The original question will appear here.";
        elements.questionsHint.textContent = "Questions hidden. Click to reveal the station brief.";
        elements.questionsBody.textContent = "The question prompt is currently hidden.";
        elements.answerHint.textContent = 'Answer hidden. Click "Show answer" to reveal the marking rubric and GPT answer.';
        elements.rubricBody.textContent = "The marking rubric is currently hidden.";
        elements.answerBody.textContent = "The answer is currently hidden.";
        elements.toggleQuestionsButton.textContent = "Show questions";
        elements.toggleAnswerButton.textContent = "Show answer";
        return;
      }}

      renderMeta(station);
      elements.stemBody.textContent = station.stem;
      elements.originalQuestionBody.textContent = station.original_question || "No original question text available.";

      if (state.questionsVisible) {{
        elements.questionsHint.textContent = "Questions visible.";
        elements.questionsBody.textContent = station.display_question || "No question text available.";
        elements.toggleQuestionsButton.textContent = "Hide questions";
      }} else {{
        elements.questionsHint.textContent = "Questions hidden. Click to reveal the station brief.";
        elements.questionsBody.textContent = "The question prompt is currently hidden.";
        elements.toggleQuestionsButton.textContent = "Show questions";
      }}

      if (state.answerVisible) {{
        elements.answerHint.textContent = "Marking rubric and GPT answer visible.";
        elements.rubricBody.textContent = station.marking_rubric || "No marking rubric is stored for this station.";
        elements.answerBody.textContent = station.answer || "No GPT answer is stored for this station.";
        elements.toggleAnswerButton.textContent = "Hide answer";
      }} else {{
        elements.answerHint.textContent = 'Answer hidden. Click "Show answer" to reveal the marking rubric and GPT answer.';
        elements.rubricBody.textContent = "The marking rubric is currently hidden.";
        elements.answerBody.textContent = "The answer is currently hidden.";
        elements.toggleAnswerButton.textContent = "Show answer";
      }}
    }}

    function setCurrentStation(station) {{
      state.currentStation = station;
      state.questionsVisible = false;
      state.answerVisible = false;
      renderPanels();
    }}

    function pickRandom(pool) {{
      if (!pool.length) {{
        window.alert("No stations match the current filters.");
        return;
      }}
      const choice = pool[Math.floor(Math.random() * pool.length)];
      setCurrentStation(choice);
    }}

    function attachEvents() {{
      [elements.yearFilter, elements.typeFilter, elements.topicFilter, elements.subtopicFilter, elements.examFilter].forEach((element) => {{
        element.addEventListener("input", () => {{
          updatePoolLabel();
          if (state.currentStation && !filteredStations().includes(state.currentStation)) {{
            state.currentStation = null;
            renderPanels();
          }}
        }});
      }});

      elements.pickFilteredButton.addEventListener("click", () => pickRandom(filteredStations()));
      elements.pickAllButton.addEventListener("click", () => pickRandom(stations));
      elements.clearFiltersButton.addEventListener("click", () => {{
        elements.yearFilter.value = "Any";
        elements.typeFilter.value = "Any";
        elements.topicFilter.value = "";
        elements.subtopicFilter.value = "";
        elements.examFilter.value = "";
        updatePoolLabel();
      }});
      elements.toggleQuestionsButton.addEventListener("click", () => {{
        if (!state.currentStation) {{
          return;
        }}
        state.questionsVisible = !state.questionsVisible;
        renderPanels();
      }});
      elements.toggleAnswerButton.addEventListener("click", () => {{
        if (!state.currentStation) {{
          return;
        }}
        state.answerVisible = !state.answerVisible;
        renderPanels();
      }});
    }}

    populateFilters();
    attachEvents();
    updatePoolLabel();
    pickRandom(stations);
  </script>
</body>
</html>
"""


def write_html(output_path: Path, html: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")


def write_github_pages_support(site_output: Path) -> None:
    site_output.parent.mkdir(parents=True, exist_ok=True)
    site_output.parent.joinpath(".nojekyll").write_text("", encoding="utf-8")


def open_in_browser(output_path: Path) -> None:
    try:
        subprocess.run(["open", str(output_path)], check=True)
    except Exception:
        print(f"Open this file in your browser: {output_path}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate and open the AFRM OSCE Picker.")
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK, help="Path to the .xlsx workbook.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Path for the generated HTML app.")
    parser.add_argument("--site-output", type=Path, default=DEFAULT_SITE_OUTPUT, help="Path for the GitHub Pages site entry point.")
    parser.add_argument("--no-open", action="store_true", help="Generate the HTML without opening it.")
    args = parser.parse_args()

    try:
        stations = load_stations(args.workbook)
        html = html_template(stations, args.workbook)
        write_html(args.output, html)
        write_html(args.site_output, html)
        write_github_pages_support(args.site_output)
    except Exception as exc:
        print(f"Failed to build AFRM OSCE Picker: {exc}", file=sys.stderr)
        return 1

    print(args.output)
    print(args.site_output)
    if not args.no_open:
        open_in_browser(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
