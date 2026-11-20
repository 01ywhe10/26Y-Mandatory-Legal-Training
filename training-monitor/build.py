"""Build the offline training monitor from an LMS export and Employee Master."""
import argparse
import base64
import json
import re
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
FIELDS = ("부서", "사번", "이름", "과정명", "진도율", "수료여부", "최종학습일", "수료일자", "학습기간")


class ReportTable(HTMLParser):
    """Read report cells as text and ignore the embedded markup."""
    def __init__(self):
        super().__init__()
        self.rows = []
        self.row = None
        self.cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell = []

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.cell is not None:
            self.row.append(" ".join("".join(self.cell).split()))
            self.cell = None
        elif tag == "tr" and self.row is not None:
            self.rows.append(self.row)
            self.row = None


def sheet_rows(path):
    with ZipFile(path) as book:
        strings = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in item.findall(".//m:t", NS))
                       for item in root.findall("m:si", NS)]
        rows = []
        root = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
        for row in root.findall(".//m:row", NS):
            values = {}
            for cell in row.findall("m:c", NS):
                col = 0
                for letter in re.match(r"[A-Z]+", cell.attrib["r"]).group():
                    col = col * 26 + ord(letter) - ord("A") + 1
                value = cell.findtext("m:v", "", NS)
                kind = cell.get("t")
                if kind == "s":
                    value = strings[int(value)]
                elif kind == "inlineStr":
                    value = "".join(t.text or "" for t in cell.findall(".//m:t", NS))
                if value.strip():
                    values[col] = value
            if values:
                rows.append([values.get(i, "") for i in range(1, max(values) + 1)])
        return rows


def load_employees(path):
    rows = sheet_rows(path)
    header = rows[0]
    emp_map = {}
    stats = {"total": 0, "active": 0, "leave": 0, "retired": 0, "excluded": 0}
    for row in rows[1:]:
        d = {header[i]: row[i] if i < len(row) else "" for i in range(len(header))}
        emp_id = d.get("사번", "").strip()
        if not emp_id:
            continue
        if d.get("조직명", "") in ("선박", "예비원") or d.get("직위", "") == "비상근고문(임원)":
            stats["excluded"] += 1
            continue
        emp_map[emp_id] = d
        stats["total"] += 1
        status = d.get("현직구분", "")
        if status == "재직":
            stats["active"] += 1
        elif status == "휴직":
            stats["leave"] += 1
        elif status == "퇴직":
            stats["retired"] += 1
    return emp_map, stats


def find_employee(emp_map, raw_id, name):
    raw_id = str(raw_id).strip()
    if raw_id in emp_map:
        return emp_map[raw_id]
    if len(raw_id) >= 7 and raw_id[-7:] in emp_map:
        return emp_map[raw_id[-7:]]
    if name == "차재승" and "K26010" in emp_map:
        return emp_map["K26010"]
    for emp in emp_map.values():
        if emp.get("성명") == name:
            return emp
    return None


def load_report(path, emp_map=None):
    rows = sheet_rows(path)
    if any(row[0].lstrip().lower().startswith(("<!doctype html", "<html")) for row in rows[:10]):
        parser = ReportTable()
        parser.feed("\n".join(" ".join(row) for row in rows))
        parser.close()
        rows = parser.rows
    header = None
    records = []
    for row in rows:
        columns = ["".join(cell.split()) for cell in row]
        if set(FIELDS).issubset(columns):
            header = {key: columns.index(key) for key in FIELDS}
            if "아이디" in columns:
                header["아이디"] = columns.index("아이디")
            continue
        if header is None or not any(row):
            continue
        def value(key):
            if key not in header:
                return ""
            index = header[key]
            return row[index].strip() if index < len(row) else ""
        status = value("수료여부")
        progress_text = value("진도율").replace("%", "").strip()
        if not value("사번") or not value("이름") or not status or not value("과정명"):
            raise ValueError("사번, 이름, 과정명 또는 수료 여부가 비어 있는 행이 있습니다.")
        progress = float(progress_text)
        if not 0 <= progress <= 100:
            raise ValueError("진도율은 0에서 100 사이여야 합니다.")
        
        raw_id = value("사번")
        name = value("이름")
        lms_account = value("아이디")
        if lms_account == "sampyo-sunghun.shin":
            raw_id = "1008821"
        emp = find_employee(emp_map, raw_id, name) if emp_map else None

        parent_dept = emp.get("상위조직명") or emp.get("조직명") or "(주)삼표시멘트" if emp else ""
        records.append({
            "id": emp.get("사번", raw_id) if emp else raw_id,
            "rawId": raw_id,
            "name": name,
            "parentDept": parent_dept,
            "department": emp.get("조직명", value("부서")) if emp else value("부서"),
            "position": emp.get("직위", "") if emp else "",
            "jobRole": emp.get("직책", "") if emp else "",
            "workplace": emp.get("근무지", "") if emp else "",
            "employmentStatus": emp.get("현직구분", "재직") if emp else "재직",
            "course": value("과정명"),
            "progress": progress,
            "status": status,
            "lastStudy": value("최종학습일").strip("-"),
            "completed": value("수료일자").strip("-"),
            "period": value("학습기간"),
        })
    if not records:
        raise ValueError("교육 현황 표 또는 교육 기록을 찾을 수 없습니다.")
    return records


def main():
    folder = Path(__file__).resolve().parent
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("xlsx", type=Path, help="LMS 교육 수료현황 엑셀 파일")
    cli.add_argument("--employees", type=Path, default=None, help="인사원장 엑셀 파일 (기본값: 자동 탐색)")
    args = cli.parse_args()

    # Find employee master file if not specified
    emp_path = args.employees
    if not emp_path:
        candidates = list(folder.parent.glob("employees_*.xlsx")) + list(folder.glob("employees_*.xlsx"))
        if candidates:
            emp_path = candidates[0]

    emp_map = {}
    stats = {}
    emp_master_date = ""
    if emp_path and emp_path.exists():
        emp_map, stats = load_employees(emp_path)
        m_date = re.search(r"(20\d{2})(\d{2})(\d{2})", emp_path.stem)
        if m_date:
            emp_master_date = f"{m_date[1]}-{m_date[2]}-{m_date[3]}"
        print(f"인사원장 로드 완료 ({emp_path.name}): 기초명단 {stats['total']:,}명 (재직 {stats['active']:,}명, 휴직 {stats['leave']}명 / 해상·임원·고문 {stats['excluded']}명 제외)")

    records = load_report(args.xlsx, emp_map)
    stamp = re.search(r"(20\d{2})(\d{2})(\d{2})(\d{2})(\d{2})\d*$", args.xlsx.stem)

    lms_ids = {r["id"] for r in records}
    non_targets = []
    if emp_map:
        exec_positions = {"회장", "수석부회장", "부회장", "사외이사", "상무", "고문(직원)"}
        for emp_id, emp in emp_map.items():
            if emp_id in lms_ids:
                continue
            org = emp.get("조직명", "")
            status_cur = emp.get("현직구분", "")
            emp_type = emp.get("고용유형", "")
            pos = emp.get("직위", "")
            if emp_type in ("임원", "겸직") or pos in exec_positions:
                stats["excluded"] += 1
                stats["total"] -= 1
                if status_cur == "재직":
                    stats["active"] -= 1
                continue
            if emp_id == "K25008" or emp.get("성명") == "김지명":
                category = "겸직"
            elif status_cur == "휴직":
                category = "휴직"
            else:
                category = "일반 재직(미등록)"
            non_targets.append({
                "id": emp_id,
                "name": emp.get("성명", ""),
                "parentDept": emp.get("상위조직명") or emp.get("조직명") or "(주)삼표시멘트",
                "department": org or "(부서 미지정)",
                "position": pos,
                "jobRole": emp.get("직책", ""),
                "employmentType": emp_type,
                "employmentStatus": status_cur,
                "workplace": emp.get("근무지", ""),
                "hireDate": emp.get("입사일", ""),
                "category": category,
            })

    payload = {
        "snapshot": f"{stamp[1]}-{stamp[2]}-{stamp[3]} {stamp[4]}:{stamp[5]}" if stamp else "",
        "empMasterDate": emp_master_date,
        "companyStats": stats,
        "records": records,
        "nonTargets": non_targets,
    }

    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    template = (folder / "dashboard.html").read_text(encoding="utf-8")

    mp4_candidates = list(folder.parent.glob("*.mp4")) + list(folder.glob("*.mp4"))
    video_path = mp4_candidates[0] if mp4_candidates else None
    if video_path and video_path.exists():
        video_b64 = base64.b64encode(video_path.read_bytes()).decode("ascii")
        template = template.replace(
            'src="./법정의무교육_안내.mp4"',
            f'src="data:video/mp4;base64,{video_b64}"',
            1,
        )
        print(f"안내 동영상 인라인 임베드 완료 ({video_path.name}: {video_path.stat().st_size / 1024 / 1024:.2f} MB)")

    output = folder.parent / "training-monitor.html"
    compiled = template.replace("<!-- REPORT_DATA -->", data, 1)
    output.write_text(compiled, encoding="utf-8")

    index_output = folder.parent / "index.html"
    index_output.write_text(compiled.replace("<body>", '<body data-default-view="lookup">', 1), encoding="utf-8")
    
    statuses = Counter(row["status"] for row in records)
    print(f"LMS 교육 기록: {len(records):,}건 · 사번 {len({row['id'] for row in records}):,}개 · 상태 {dict(statuses)}")
    matched_count = sum(bool(r.get("parentDept")) for r in records)
    print(f"인사원장 매칭: {matched_count:,}건 ({matched_count/len(records)*100:.1f}%) · 부서 미기재 0건 해소 완료")
    print(f"생성 완료: {output}")


if __name__ == "__main__":
    main()
