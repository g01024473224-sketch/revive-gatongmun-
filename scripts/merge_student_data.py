"""
학생별 데이터 병합기: 매쓰플랫 평가 + 진도계획표 진도 → 통합 가통문 데이터

핵심 규칙:
- 학년은 매쓰플랫이 우선 (진도계획표와 다르면 매쓰플랫 사용 + 경고)
- 전국 평균 = 같은 학년 + 같은 학습지를 푼 학생들의 평균 점수
- 재시 이력은 원본 + 재시/재재시... 모두 포함
- 매쓰플랫에만 있음 → 경고 (진도 없이 평가만)
- 진도계획표에만 있음 → 평가 섹션 생략 (진도만)
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

DEFAULT_AVG_FALLBACK = 75.0  # 반 평균 계산 불가 시 기본값
MAX_DISPLAYED_AVG = 85.0  # 90% 이상이면 이 값으로 고정 (원장님 요청)

# 난이도 → 라벨 매핑
DIFFICULTY_LABELS = {
    1: "연산",
    2: "기본",
    3: "유형 학습",
    4: "심화",
    5: "고난도",
}

# 반 강제 매핑 (진도계획표에 없어도 이 반으로 분류)
# (학교급, 학년, 이름) → 학생 코드
FORCED_CLASS = {
    ("중학교", "3", "박태상"): "중3S2",
    ("중학교", "3", "박관우"): "중3S2",
    ("중학교", "3", "이정원"): "중3S2",
    ("중학교", "3", "주소연"): "중3S2",
}

# 반 단위 진도 공유 (해당 반의 진도가 없는 학생은 대표 학생 진도 사용)
# 학생 코드 → 진도를 가져올 대표 학생 이름
SHARED_PROGRESS_REPRESENTATIVE = {
    "중3S2": "박관우",
}


def load_mathflat_all(output_dir: Path) -> dict:
    """output/*_parsed.json 모두 로드 → 학년별 구조로 반환."""
    combined = {"by_grade_name": {}, "raw": []}
    for jf in sorted(output_dir.glob("*_parsed.json")):
        # 진도계획표_전체.parsed.json 같은 것은 제외
        if "진도계획표" in jf.name or jf.name.startswith("progress_"):
            continue
        with jf.open(encoding="utf-8") as f:
            data = json.load(f)
        grade = data.get("grade")
        level = data.get("school_level")
        for student_name, evals in data.get("students", {}).items():
            key = (level, grade, student_name)
            combined["by_grade_name"][key] = evals
            combined["raw"].append({
                "school_level": level,
                "grade": grade,
                "name": student_name,
                "evals": evals,
            })
    return combined


def load_progress(progress_json: Path) -> dict:
    """progress_YYYY_MM.json 로드."""
    with progress_json.open(encoding="utf-8") as f:
        return json.load(f)


def grade_to_level_grade(code: str) -> tuple[str, int] | None:
    """학생코드에서 학교급+학년 추출. '중1U2' → ('중학교', 1)."""
    if not code:
        return None
    import re
    m = re.match(r"^(초|중|고)(\d)", code)
    if not m:
        return None
    level_map = {"초": "초등학교", "중": "중학교", "고": "고등학교"}
    return (level_map[m.group(1)], int(m.group(2)))


def compute_national_averages(all_mathflat: list[dict]) -> dict:
    """
    학년별 × 학습지명별 평균 점수 계산.
    키: (school_level, grade, 학습지명)
    값: 평균 점수 (float)
    """
    # 같은 학습지 원본(재시차수=0)만 모아서 평균 계산
    # 재시는 평균 왜곡하므로 제외 (가통문 샘플에서도 원본 기준으로 보임)
    scores_by_key: dict = defaultdict(list)
    for rec in all_mathflat:
        level = rec["school_level"]
        grade = rec["grade"]
        name = rec["name"]
        evals = rec["evals"]
        for kind in ("weekly_evals", "unit_evals"):
            for e in evals.get(kind, []):
                if e.get("재시차수", 0) == 0 and e.get("점수") is not None:
                    key = (level, grade, e["단원명"])
                    scores_by_key[key].append(e["점수"])

    result = {}
    for k, v in scores_by_key.items():
        if not v:
            continue
        avg = round(sum(v) / len(v))
        # 원장님 요청: 90% 이상은 85%로 고정 (학부모 기대치 관리)
        if avg >= 90:
            avg = int(MAX_DISPLAYED_AVG)
        result[k] = avg
    return result


def detect_test_prep_done(activities: list[dict]) -> bool:
    """
    매쓰플랫 학습지 데이터에서 시험 대비를 '실제로 한' 흔적 찾기.
    → 가통문에 "시험 대비를 했습니다" (과거형)
    """
    keywords = ["기출", "내신", "중간고사", "기말고사", "시험 대비", "시험대비"]
    for a in activities:
        name = str(a.get("원본_학습지명") or a.get("단원명") or "")
        tag = str(a.get("태그") or "")
        for kw in keywords:
            if kw in name or kw in tag:
                return True
    return False


def detect_pending_retests(weekly_evals: list[dict], unit_evals: list[dict]) -> list[dict]:
    """
    70점 미만인데 아직 재시를 치지 못한 학습지 찾기.
    같은 학습지(단원명+주차)에 재시차수 0인 점수가 70 미만이고,
    재시차수 >= 1인 기록이 없으면 '재시 예정'으로 판단.
    """
    def key_of(e):
        # 주간평가는 주차로 구분, 선행평가는 단원명만
        return (e.get("주차"), e["단원명"])

    pending = []
    for evals in (weekly_evals, unit_evals):
        # 학습지별 그룹핑
        groups: dict = defaultdict(list)
        for e in evals:
            groups[key_of(e)].append(e)

        for key, records in groups.items():
            # 원본(재시차수 0) 점수가 70 미만인 것만
            originals = [r for r in records if r.get("재시차수", 0) == 0]
            if not originals:
                continue
            last_original = max(originals, key=lambda r: r.get("날짜") or "")
            if last_original.get("점수") is None or last_original["점수"] >= 70:
                continue
            # 최종 재시 차수 확인: 재시 중 가장 높은 차수가 70 이상이면 통과, 아니면 재시 필요
            latest_retry = max(records, key=lambda r: r.get("재시차수", 0))
            if latest_retry.get("재시차수", 0) >= 1 and latest_retry.get("점수") and latest_retry["점수"] >= 70:
                continue  # 재시 통과
            pending.append({
                "단원명": last_original["단원명"],
                "주차": last_original.get("주차"),
                "점수": last_original["점수"],
                "최종재시차수": latest_retry.get("재시차수", 0),
            })
    return pending


def compute_student_analytics(activities: list[dict]) -> dict:
    """
    학생의 모든 학습지 활동을 분석.
    반환: {
      "total_count": 총 학습지 수,
      "total_problems": 총 문항 수,
      "by_difficulty": {1: {...}, 2: {...}, ...},
      "strengths": [상위 단원들],
      "weaknesses": [하위 단원들],
    }
    """
    # 원본만 (재시 제외, 중복 점수 왜곡 방지)
    base = [a for a in activities if a.get("재시차수", 0) == 0 and a.get("점수") is not None]

    # 난이도별 집계
    by_diff: dict = defaultdict(lambda: {"count": 0, "scores": [], "problems": 0, "단원": set()})
    for a in base:
        d = a.get("난이도")
        if d is None:
            continue
        by_diff[d]["count"] += 1
        by_diff[d]["scores"].append(a["점수"])
        by_diff[d]["problems"] += a.get("문항수", 0)
        by_diff[d]["단원"].add(a["단원명"])

    difficulty_summary = {}
    for d, info in by_diff.items():
        scores = info["scores"]
        difficulty_summary[d] = {
            "label": DIFFICULTY_LABELS.get(d, f"난이도{d}"),
            "count": info["count"],
            "avg_score": round(sum(scores) / len(scores)) if scores else None,
            "total_problems": info["problems"],
            "unit_count": len(info["단원"]),
        }

    # 단원별 평균 (같은 단원 여러 건 있으면 평균)
    unit_scores: dict = defaultdict(list)
    for a in base:
        unit_scores[a["단원명"]].append(a["점수"])

    unit_avg = [
        {"단원": u, "avg_score": round(sum(s) / len(s)), "count": len(s)}
        for u, s in unit_scores.items()
    ]
    # 점수 상위/하위 (학습지 명 중에서 "주차" 식이 아닌 실제 단원명만)
    # 주간 TEST 제목은 "03월 2주 (중1)" 같은 형식이라 제외
    is_real_unit = lambda u: not re.match(r"^\d+월\s*\d+주", u)
    real_units = [x for x in unit_avg if is_real_unit(x["단원"])]
    real_units.sort(key=lambda x: x["avg_score"], reverse=True)

    strengths = [x for x in real_units if x["avg_score"] >= 85][:5]
    weaknesses = [x for x in real_units if x["avg_score"] < 70][:5]

    return {
        "total_count": len(base),
        "total_problems": sum(a.get("문항수", 0) for a in base),
        "by_difficulty": difficulty_summary,
        "strengths": strengths,
        "weaknesses": weaknesses,
    }


def build_student_records(
    all_mathflat: dict,
    progress_data: dict,
) -> list[dict]:
    """
    매쓰플랫 + 진도계획표를 병합하여 학생별 통합 레코드 생성.
    """
    import re  # local import for regex usage
    globals()["re"] = re  # compute_student_analytics에서도 쓰기 위함
    records: list[dict] = []

    # 매쓰플랫 학생 키: (학교급, 학년, 이름)
    mathflat_by_key = all_mathflat["by_grade_name"]

    # 진도계획표 학생 인덱스: 이름 → [학생정보...]
    progress_by_name: dict[str, list[dict]] = defaultdict(list)
    # 진도계획표 학생 인덱스: 코드 → [학생정보...]  (반 공유용)
    progress_by_code: dict[str, list[dict]] = defaultdict(list)
    for s in progress_data.get("students", []):
        progress_by_name[s["name"]].append(s)
        progress_by_code[s["code"]].append(s)

    # 전국 평균 테이블
    avg_table = compute_national_averages(all_mathflat["raw"])

    # 매쓰플랫 기준으로 순회 (매쓰플랫 학년 우선)
    used_progress_keys = set()

    for (level, grade, name), evals in mathflat_by_key.items():
        warnings: list[str] = []

        # 진도 데이터 매칭
        progress_candidates = progress_by_name.get(name, [])
        progress = None

        # 학년 일치 먼저 시도
        for p in progress_candidates:
            p_lg = grade_to_level_grade(p["code"])
            if p_lg and p_lg == (level, int(grade)):
                progress = p
                break

        # 학년 불일치 — 유일한 후보라면 사용하고 경고
        if progress is None and len(progress_candidates) == 1:
            p = progress_candidates[0]
            p_lg = grade_to_level_grade(p["code"])
            if p_lg:
                warnings.append(
                    f"진도계획표의 학생코드 '{p['code']}'와 매쓰플랫의 학년 '{level} {grade}'이 불일치. "
                    f"매쓰플랫 기준으로 생성."
                )
            progress = p

        # 강제 반 매핑 확인: 진도 없는 학생이지만 공유 반이면 대표 학생의 진도 사용
        forced_code = FORCED_CLASS.get((level, str(grade), name))
        if progress is None and forced_code:
            rep_name = SHARED_PROGRESS_REPRESENTATIVE.get(forced_code)
            if rep_name:
                for p in progress_by_code.get(forced_code, []):
                    if p["name"] == rep_name:
                        progress = p  # 대표 학생의 진도를 공유
                        break
                # 대표 찾지 못하면 같은 반 아무 진도라도 사용
                if progress is None and progress_by_code.get(forced_code):
                    progress = progress_by_code[forced_code][0]

        # 아예 매칭 실패 — 매쓰플랫에만 존재
        if progress is None and not forced_code:
            warnings.append("진도계획표에 학생이 없음. 평가 섹션만 생성됨.")

        if progress is not None:
            used_progress_keys.add((progress["code"], progress["name"]))

        # 평가 정리 (재시 포함, 단원별/주차별 정렬)
        weekly = sorted(
            evals.get("weekly_evals", []),
            key=lambda e: (e.get("주차") or 99, e.get("재시차수", 0), e.get("날짜") or ""),
        )
        unit = sorted(
            evals.get("unit_evals", []),
            key=lambda e: (e.get("단원명") or "", e.get("재시차수", 0), e.get("날짜") or ""),
        )

        # 각 평가에 전국 평균 붙이기 (90% 이상은 85%로 고정 — MAX_DISPLAYED_AVG)
        for lst in (weekly, unit):
            for e in lst:
                key = (level, grade, e["단원명"])
                avg = avg_table.get(key, DEFAULT_AVG_FALLBACK)
                if avg >= 90:
                    avg = int(MAX_DISPLAYED_AVG)
                e["전국평균"] = avg

        # 학생 코드: 강제 매핑 > 진도계획표 > 학년만
        if forced_code:
            code = forced_code
        elif progress:
            code = progress["code"]
        else:
            code = f"{'초중고'['초중고'.index(level[0])]}{grade}"

        # 분석 데이터 (all_activities 기반)
        analytics = compute_student_analytics(evals.get("all_activities", []))

        # 재시 예정 감지 (70점 미만인데 재시 안 친 학습지)
        pending_retests = detect_pending_retests(weekly, unit)

        # 시험대비 모드 감지 (두 가지 구분)
        # (1) plan: 진도계획표에 "시험대비" → 이번 달에 시험을 치를 계획 ("~할 예정")
        test_prep_plan = False
        if progress:
            for role in progress.get("roles", []):
                all_content = list(role.get("주차별", {}).values()) + [role.get("교재", "")]
                for c in all_content:
                    if "시험" in str(c) and "대비" in str(c):
                        test_prep_plan = True
                        break

        # (2) done: 매쓰플랫에 시험 대비 흔적 → 이미 진행 ("~했습니다")
        test_prep_done = detect_test_prep_done(evals.get("all_activities", []))

        records.append({
            "code": code,
            "name": name,
            "school_level": level,
            "grade": grade,
            "progress": progress,  # None or full progress dict
            "weekly_evals": weekly,
            "unit_evals": unit,
            "analytics": analytics,
            "pending_retests": pending_retests,
            "test_prep_plan": test_prep_plan,  # 앞으로 할 계획 (가통문 배포 시점 기준 미래)
            "test_prep_done": test_prep_done,  # 이미 한 활동 (가통문 배포 시점 기준 과거)
            "warnings": warnings,
            "has_evaluations": bool(weekly or unit),
            "has_progress": progress is not None,
        })

    # 진도계획표에만 있는 학생 (매쓰플랫에 평가 없음)
    for name, candidates in progress_by_name.items():
        for p in candidates:
            key = (p["code"], p["name"])
            if key in used_progress_keys:
                continue
            # 이미 매쓰플랫과 매칭된 다른 코드의 동명 학생이 있으면 스킵 판단
            if any(name == r["name"] for r in records):
                # 동명 매쓰플랫 학생이 다른 학년에 존재 — 별개로 추가
                pass

            lg = grade_to_level_grade(p["code"])
            if not lg:
                continue
            level, grade = lg
            records.append({
                "code": p["code"],
                "name": p["name"],
                "school_level": level,
                "grade": str(grade),
                "progress": p,
                "weekly_evals": [],
                "unit_evals": [],
                "analytics": compute_student_analytics([]),
                "warnings": [],  # 평가 없음 → 레이아웃 B (큰 달력)
                "has_evaluations": False,
                "has_progress": True,
            })

    # 정렬: 학교급 → 학년 → 학생코드
    level_order = {"초등학교": 0, "중학교": 1, "고등학교": 2}
    records.sort(key=lambda r: (
        level_order.get(r["school_level"], 9),
        int(r["grade"]) if str(r["grade"]).isdigit() else 9,
        r["code"], r["name"],
    ))

    return records


def main():
    if len(sys.argv) < 3:
        print("사용법: python merge_student_data.py <연도> <월>")
        print("예시:   python merge_student_data.py 2026 4")
        sys.exit(1)

    year = int(sys.argv[1])
    month = int(sys.argv[2])

    base = Path(__file__).parent.parent
    output_dir = base / "output"
    progress_json = output_dir / f"progress_{year}_{month:02d}.json"

    if not progress_json.exists():
        print(f"❌ {progress_json.name} 없음. 먼저 parse_progress.py 실행.")
        sys.exit(1)

    mf = load_mathflat_all(output_dir)
    pg = load_progress(progress_json)
    records = build_student_records(mf, pg)

    out_path = output_dir / f"merged_{year}_{month:02d}.json"
    with out_path.open("w", encoding="utf-8") as f:
        json.dump({
            "year": year,
            "month": month,
            "students": records,
        }, f, ensure_ascii=False, indent=2)

    # 리포트
    n_total = len(records)
    n_both = sum(1 for r in records if r["has_evaluations"] and r["has_progress"])
    n_only_math = sum(1 for r in records if r["has_evaluations"] and not r["has_progress"])
    n_only_prog = sum(1 for r in records if not r["has_evaluations"] and r["has_progress"])

    print(f"✅ 병합 완료: {year}년 {month}월")
    print(f"   전체 학생: {n_total}명")
    print(f"     ├─ 완전 (매쓰플랫 + 진도계획표 ): {n_both}명")
    print(f"     ├─ 매쓰플랫만 있음 ⚠️: {n_only_math}명")
    print(f"     └─ 진도계획표만 있음 (평가 없음): {n_only_prog}명")
    print()
    print("   ⚠️ 경고 대상 학생:")
    has_warn = False
    for r in records:
        if r["warnings"]:
            has_warn = True
            for w in r["warnings"]:
                print(f"     · {r['code']} {r['name']}: {w}")
    if not has_warn:
        print("     (없음)")

    print()
    print("   👤 샘플 학생 (상위 5명):")
    for r in records[:5]:
        n_w = len(r["weekly_evals"])
        n_u = len(r["unit_evals"])
        roles = "+".join(x["role"] for x in (r["progress"]["roles"] if r["progress"] else []))
        print(f"     {r['code']:8s} {r['name']:5s} | 주간평가 {n_w}건, 선행평가 {n_u}건 | 진도({roles or 'X'})")

    print(f"\n   출력: {out_path}")


if __name__ == "__main__":
    main()
