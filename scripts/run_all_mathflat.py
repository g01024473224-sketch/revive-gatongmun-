"""모든 학년 매쓰플랫 Excel 일괄 파싱 + 요약 리포트."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from parse_mathflat import parse_mathflat_excel

BASE = Path(__file__).parent.parent
OUTPUT = BASE / "output"
OUTPUT.mkdir(exist_ok=True)


def short_grade(filename: str) -> str:
    """파일명에서 학년 키 추출 (중1, 초6 등)."""
    import re
    m = re.search(r"(초\d|중\d)", filename)
    return m.group(1) if m else filename


def main():
    excel_files = sorted(BASE.glob("*학습지 학습내역*.xlsx"))
    if not excel_files:
        print("⚠️  매쓰플랫 Excel 파일을 찾을 수 없습니다.")
        return

    all_data = {}
    retry_highlights = []  # 재시 이력이 있는 학생 모음

    for xlsx in excel_files:
        grade = short_grade(xlsx.name)
        data = parse_mathflat_excel(xlsx)

        output_path = OUTPUT / f"{grade}_parsed.json"
        with output_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        all_data[grade] = data

        # 통계
        n_students = len(data["students"])
        n_weekly = sum(len(s["weekly_evals"]) for s in data["students"].values())
        n_unit = sum(len(s["unit_evals"]) for s in data["students"].values())

        # 재시 학생 + 70점 미만
        for name, s in data["students"].items():
            for evals in (s["weekly_evals"], s["unit_evals"]):
                for e in evals:
                    if e["재시차수"] > 0 or (e["점수"] is not None and e["점수"] < 70):
                        retry_highlights.append({
                            "학년": grade, "학생": name,
                            "단원": e["단원명"], "재시차수": e["재시차수"],
                            "점수": e["점수"], "주차": e.get("주차"),
                        })

        print(f"[{grade}] 학생 {n_students}명 | 주간평가 {n_weekly}건 | 선행평가 {n_unit}건")

    # 재시/미달 하이라이트
    print("\n===== 재시 또는 70점 미만 기록 (선생님 주의 필요) =====")
    for h in retry_highlights:
        retry_tag = f"재시{h['재시차수']}차" if h["재시차수"] > 0 else "원본"
        week_tag = f"주{h['주차']}" if h.get("주차") else "   "
        red = "🔴" if (h["점수"] is not None and h["점수"] < 70) else "  "
        print(f"  {red} {h['학년']} {h['학생']:6s} | {week_tag} | {retry_tag:6s} | "
              f"{h['점수']}점 | {h['단원']}")

    # 전체 요약
    all_students = sum(len(d["students"]) for d in all_data.values())
    all_weekly = sum(
        len(s["weekly_evals"])
        for d in all_data.values() for s in d["students"].values()
    )
    all_unit = sum(
        len(s["unit_evals"])
        for d in all_data.values() for s in d["students"].values()
    )
    print(f"\n===== 전체 요약 =====")
    print(f"총 학생: {all_students}명")
    print(f"총 주간평가: {all_weekly}건")
    print(f"총 선행평가: {all_unit}건")
    print(f"재시/미달 기록: {len(retry_highlights)}건")


if __name__ == "__main__":
    main()
