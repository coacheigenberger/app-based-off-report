
from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict
from typing import Any
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


@dataclass
class GameResult:
    week: str = ""
    opponent: str = ""
    opponent_record: str = ""
    score: str = ""
    result: str = ""  # W/L/T


@dataclass
class OpponentOverview:
    record: str = ""
    sacks: str = ""
    interceptions: str = ""
    fumble_recoveries: str = ""
    source_url: str = ""
    games: list[GameResult] | None = None
    last_five_record: str = ""
    average_margin: str = ""
    one_possession_games: str = ""
    wins_vs_winning_records: str = ""

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["games"] = [asdict(g) for g in (self.games or [])]
        return d


def _normalize(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _all_json_nodes(obj: Any):
    yield obj
    if isinstance(obj, dict):
        for v in obj.values():
            yield from _all_json_nodes(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _all_json_nodes(v)


def _flatten_text_values(obj: Any) -> str:
    vals = []
    for n in _all_json_nodes(obj):
        if isinstance(n, (str, int, float)):
            vals.append(str(n))
    return " ".join(vals)


def _safe_get(url: str) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; DefenseAnalyst/1.0; +https://openai.com)",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    return r.text


def _extract_embedded_json(soup: BeautifulSoup) -> list[Any]:
    blobs = []
    for script in soup.find_all("script"):
        txt = script.string or script.get_text("", strip=True)
        if not txt:
            continue
        if script.get("id") == "__NEXT_DATA__":
            try:
                blobs.append(json.loads(txt))
            except Exception:
                pass
        elif "{" in txt and "}" in txt:
            # Best effort: pull JSON-looking assignment payloads.
            for m in re.finditer(r"(\{.*\}|\[.*\])", txt, flags=re.S):
                raw = m.group(1)
                try:
                    blobs.append(json.loads(raw))
                    break
                except Exception:
                    continue
    return blobs


def _find_team_link(list_url: str, soup: BeautifulSoup, opponent: str) -> str | None:
    target = opponent.lower().strip()
    if not target:
        return None
    best = None
    for a in soup.find_all("a", href=True):
        text = _normalize(a.get_text(" "))
        href = a["href"]
        combined = f"{text} {href}".lower()
        if target in combined and "/teams" in href:
            best = urljoin(list_url, href)
            # Prefer specific team pages over the generic listing.
            if best.rstrip("/") != list_url.rstrip("/"):
                return best
    return best


def _record_from_text(text: str) -> str:
    patterns = [
        r"(?:Overall\s+)?Record\s*[:\-]?\s*(\d+)\s*[-–]\s*(\d+)",
        r"Overall\s*[:\-]?\s*(\d+)\s*[-–]\s*(\d+)",
        r"\b(\d+)\s*[-–]\s*(\d+)\s+Overall\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
    return ""


def _stat_from_text(text: str, labels: list[str]) -> str:
    for label in labels:
        pat = rf"{label}\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)"
        m = re.search(pat, text, re.I)
        if m:
            return m.group(1)
    return ""


def _stat_from_json(json_blobs: list[Any], key_terms: list[str]) -> str:
    key_terms_l = [k.lower() for k in key_terms]
    for blob in json_blobs:
        for n in _all_json_nodes(blob):
            if isinstance(n, dict):
                for k, v in n.items():
                    lk = str(k).lower()
                    if all(term in lk for term in key_terms_l) and isinstance(v, (int, float, str)):
                        sv = str(v)
                        if re.fullmatch(r"\d+(\.\d+)?", sv):
                            return sv
    return ""


def _record_from_json(json_blobs: list[Any]) -> str:
    for blob in json_blobs:
        for n in _all_json_nodes(blob):
            if isinstance(n, dict):
                keys = {str(k).lower(): k for k in n.keys()}
                wins_key = next((keys[k] for k in keys if k in {"wins", "win", "overallwins", "overall_wins"} or ("win" in k and "pct" not in k)), None)
                losses_key = next((keys[k] for k in keys if k in {"losses", "loss", "overalllosses", "overall_losses"} or "loss" in k), None)
                if wins_key and losses_key:
                    try:
                        return f"{int(n[wins_key])}-{int(n[losses_key])}"
                    except Exception:
                        pass
    return ""


def _derive_result_and_margin(score: str, team_first: bool = True) -> tuple[str, int | None]:
    nums = [int(x) for x in re.findall(r"\d+", score or "")]
    if len(nums) < 2:
        return "", None
    margin = nums[0] - nums[1] if team_first else nums[1] - nums[0]
    if margin > 0:
        return "W", margin
    if margin < 0:
        return "L", margin
    return "T", 0


def _parse_games_from_tables(soup: BeautifulSoup) -> list[GameResult]:
    games: list[GameResult] = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue
        headers = [_normalize(h.get_text(" ")).lower() for h in rows[0].find_all(["th", "td"])]
        if not headers:
            continue
        header_text = " ".join(headers)
        if not any(x in header_text for x in ["opponent", "score", "result", "game"]):
            continue
        for tr in rows[1:]:
            cells = [_normalize(td.get_text(" ")) for td in tr.find_all(["td", "th"])]
            if len(cells) < 2:
                continue
            row = dict(zip(headers, cells))
            opp = next((row[h] for h in headers if "opponent" in h), "")
            score = next((row[h] for h in headers if "score" in h), "")
            result = next((row[h] for h in headers if "result" in h or h in {"w/l", "w-l"}), "")
            week = next((row[h] for h in headers if "week" in h or "date" in h or "game" in h), "")
            opp_record = next((row[h] for h in headers if "record" in h), "")
            if not result and score:
                result, _ = _derive_result_and_margin(score)
            if opp or score:
                games.append(GameResult(week=week, opponent=opp, opponent_record=opp_record, score=score, result=result[:1].upper()))
    return games[:15]


def _parse_games_from_text(text: str) -> list[GameResult]:
    games: list[GameResult] = []
    # Best effort pattern for lines containing W/L and a score.
    for line in re.split(r"[\n\r]+", text):
        line = _normalize(line)
        if not re.search(r"\b[WLT]\b", line, re.I) or not re.search(r"\d+\s*[-–]\s*\d+", line):
            continue
        mscore = re.search(r"(\d+\s*[-–]\s*\d+)", line)
        mres = re.search(r"\b([WLT])\b", line, re.I)
        if not mscore or not mres:
            continue
        before = line[:mres.start()].strip(" -–|")
        after = line[mres.end():].strip(" -–|")
        opponent = before or after.replace(mscore.group(1), "").strip(" -–|")
        games.append(GameResult(opponent=opponent, score=mscore.group(1), result=mres.group(1).upper()))
    return games[:15]


def _parse_games_from_json(json_blobs: list[Any]) -> list[GameResult]:
    out: list[GameResult] = []
    for blob in json_blobs:
        for n in _all_json_nodes(blob):
            if not isinstance(n, dict):
                continue
            keys = {str(k).lower(): k for k in n.keys()}
            has_score = any("score" in k or k in {"homepoints", "awaypoints", "home_score", "away_score"} for k in keys)
            has_opp = any("opponent" in k or "opponentname" in k for k in keys)
            if not (has_score and has_opp):
                continue
            opp_key = next((keys[k] for k in keys if "opponent" in k), None)
            opp = _normalize(n.get(opp_key, "")) if opp_key else ""
            score = ""
            result = ""
            if "score" in keys:
                score = _normalize(n.get(keys["score"], ""))
            else:
                nums = []
                for k in ["team_score", "teamscore", "scorefor", "pointsfor", "homepoints", "awaypoints", "home_score", "away_score"]:
                    if k in keys:
                        nums.append(str(n.get(keys[k], "")))
                if len(nums) >= 2:
                    score = f"{nums[0]}-{nums[1]}"
            for rk in ["result", "outcome", "wl", "winloss"]:
                if rk in keys:
                    result = _normalize(n.get(keys[rk], ""))[:1].upper()
                    break
            if not result and score:
                result, _ = _derive_result_and_margin(score)
            rec_key = next((keys[k] for k in keys if "opponent" in k and "record" in k), None)
            opp_record = _normalize(n.get(rec_key, "")) if rec_key else ""
            week_key = next((keys[k] for k in keys if k in {"week", "date", "game", "game_number", "gamenumber"}), None)
            week = _normalize(n.get(week_key, "")) if week_key else ""
            if opp or score:
                out.append(GameResult(week=week, opponent=opp, opponent_record=opp_record, score=score, result=result))
    # Deduplicate preserving order.
    seen = set()
    dedup = []
    for g in out:
        key = (g.week, g.opponent, g.score)
        if key not in seen:
            seen.add(key)
            dedup.append(g)
    return dedup[:15]


def _compute_enhancements(games: list[GameResult]) -> dict[str, str]:
    scored = []
    for g in games:
        result, margin = _derive_result_and_margin(g.score)
        if g.result:
            result = g.result[:1].upper()
        if margin is not None:
            scored.append((result, margin, g))
    if not scored:
        return {"last_five_record": "", "average_margin": "", "one_possession_games": "", "wins_vs_winning_records": ""}

    last5 = scored[-5:]
    w = sum(1 for r, _, _ in last5 if r == "W")
    l = sum(1 for r, _, _ in last5 if r == "L")
    avg = sum(m for _, m, _ in scored) / len(scored)
    one_score = sum(1 for _, m, _ in scored if abs(m) <= 8)

    wins_vs_winning = 0
    total_vs_winning = 0
    for result, _, g in scored:
        m = re.search(r"(\d+)\s*[-–]\s*(\d+)", g.opponent_record or "")
        if m:
            ow, ol = int(m.group(1)), int(m.group(2))
            if ow > ol:
                total_vs_winning += 1
                if result == "W":
                    wins_vs_winning += 1
    wwr = f"{wins_vs_winning}-{total_vs_winning - wins_vs_winning}" if total_vs_winning else ""
    return {
        "last_five_record": f"{w}-{l}",
        "average_margin": f"{avg:+.1f}",
        "one_possession_games": str(one_score),
        "wins_vs_winning_records": wwr,
    }


def fetch_gobound_overview(url: str, opponent: str = "") -> OpponentOverview:
    """
    Best-effort GoBound scraper. It works with either a specific team page URL or the
    GoBound teams listing page plus an opponent name. If GoBound changes their markup,
    the app will still run and the overview slide will show whatever fields could be found.
    """
    html = _safe_get(url)
    soup = BeautifulSoup(html, "html.parser")

    team_url = _find_team_link(url, soup, opponent)
    if team_url and team_url.rstrip("/") != url.rstrip("/"):
        html = _safe_get(team_url)
        soup = BeautifulSoup(html, "html.parser")
        url = team_url

    text = soup.get_text("\n", strip=True)
    json_blobs = _extract_embedded_json(soup)
    json_text = " ".join(_flatten_text_values(b) for b in json_blobs)

    record = _record_from_json(json_blobs) or _record_from_text(text) or _record_from_text(json_text)
    sacks = (
        _stat_from_json(json_blobs, ["sack"]) or
        _stat_from_text(text, ["Total Sacks", "Sacks"])
    )
    interceptions = (
        _stat_from_json(json_blobs, ["interception"]) or
        _stat_from_json(json_blobs, ["int"]) or
        _stat_from_text(text, ["Total INTs", "Interceptions", "INTs"])
    )
    fumbles = (
        _stat_from_json(json_blobs, ["fumble", "recover"]) or
        _stat_from_text(text, ["Total Fumble Recoveries", "Fumble Recoveries", "Fumbles Recovered"])
    )

    games = _parse_games_from_tables(soup) or _parse_games_from_json(json_blobs) or _parse_games_from_text(text)
    enh = _compute_enhancements(games)

    return OpponentOverview(
        record=record,
        sacks=sacks,
        interceptions=interceptions,
        fumble_recoveries=fumbles,
        source_url=url,
        games=games,
        **enh,
    )
