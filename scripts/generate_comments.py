"""
학생별 AI 맞춤 코멘트 생성기

- 입력: merged_YYYY_MM.json
- 출력: merged_YYYY_MM.json (ai_comment 필드 추가)

특징:
- Claude Sonnet 4.6 사용 (가성비 우수)
- Prompt caching으로 시스템 프롬프트 재사용 (비용 90% 절감)
- 재시 이력, 점수, 진도 정보를 종합한 3~4문장 코멘트
"""

from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

from anthropic import Anthropic
from dotenv import load_dotenv

BASE = Path(__file__).parent.parent
load_dotenv(BASE / ".env")

MODEL = "claude-sonnet-4-5-20250929"  # 가성비 우수한 Sonnet 최신
MAX_TOKENS = 700


SYSTEM_PROMPT = """당신은 리바이브수학전문학원의 수학 선생님으로, 학부모님께 보내는 월간 가정통신문에 학생별 맞춤 코멘트를 작성하는 역할을 맡고 있습니다.

[난이도 체계 참고]
- 난이도 1 = 연산 (기초 계산)
- 난이도 2 = 기본 (개념 확인)
- 난이도 3 = 유형 학습 (문제 유형 숙달)
- 난이도 4 = 심화 (응용)
- 난이도 5 = 고난도

[작성 내용 구성] — 아래 요소들을 자연스럽게 녹여서 5~7문장으로 작성:
1. 이번 달 **학습 분량 언급** — 반드시 "교재 외 학습지 N문항을 풀었습니다" 식으로 표현 (학원에서 교재와 별개로 학습지도 내줌)
2. **잘한 단원 칭찬** (점수 85점 이상 단원) — 구체적 단원명
3. **부족한 단원 보충 계획** (70점 미만 단원) — 구체적 단원명
4. **재시 이력이 있으면** 극복 과정 긍정적으로 표현
5. **재시 예정(아직 재시 못 친) 학습지가 있으면** → 반드시 "오는 달에 재시를 치를 예정입니다"라는 문구 포함
6. **난이도별 학습 패턴** 코멘트 — 예:
   - 난이도 1(연산) 비중 높음 → "연산 숙달을 위해 반복 훈련을 진행했습니다"
   - 난이도 2(기본) 중심 → "개념을 탄탄히 다지는 단계에 집중했습니다"
   - 난이도 3(유형) 중심 → "다양한 문제 유형을 경험하며 실전 감각을 키웠습니다"
   - 난이도 4(심화) 포함 → "응용 능력을 키우기 위해 심화 문제에 도전했습니다"
   - 난이도 5(고난도) 포함 → "고난도 문제로 사고력을 확장했습니다"
7. 다음 달 지도 방향
8. **시험대비 관련 문구 (시제에 주의!)**:

   [※※ 가통문은 해당 월 시작 전에 학부모에게 전달됩니다. 시제를 정확히 맞추세요. ※※]

   A) **[시험대비 예정]** 플래그가 있으면 → **미래형**으로 작성:
      - "이번 달에는 시험 대비를 충실히 진행할 계획입니다"
      - "주변 학교들의 기출 문제를 풀며 실전 감각을 기를 예정입니다"
      - "지금까지 학습한 유형 중 취약 부분을 선별하여 약점을 보완해 나가겠습니다"
      - "서술형 문제 준비까지 철저히 진행할 계획입니다"
      - "내신 대비 시험으로 실전 감각을 배양하겠습니다"

   B) **[시험대비 완료]** 플래그가 있으면 → **과거형**으로 작성:
      - "시험 대비를 충실히 진행했습니다"
      - "주변 학교들의 기출 문제를 풀이하며 실전 감각을 배양했습니다"
      - "지금까지 학습한 유형 중 취약 부분을 선별하여 약점을 보완했습니다"
      - "서술형 문제 준비까지 철저히 진행했습니다"
      - "내신 대비 시험을 통해 실전 감각을 키웠습니다"

   C) 두 플래그 모두 있으면 과거 진행된 부분은 과거형, 앞으로 할 부분은 미래형으로 자연스럽게 혼합

[표현 규칙]
- 월 표기는 앞에 0을 붙이지 않음: "3월" / "4월" (O), "03월" / "04월" (X)
- "~습니다" 체 격식 있는 한국어
- 학원 원장의 시선에서 학부모에게 전달하는 톤 (따뜻하되 프로페셔널)
- 출력은 오직 코멘트 본문만 (인사말, 머리말, 맺음말 "이상입니다" 등 제외)
- 70점 미만 성적은 "어려움을 겪었으나" / "다소 부족한 부분이 있어" 식으로 부드럽게
- 재시 통과는 "회복했습니다" / "극복해냈습니다"로 긍정적으로

[피해야 할 것]
- "안녕하세요", "어머님께" 같은 인사말
- 학생 이름을 여러 번 반복
- "열심히 했습니다" 같은 뻔한 표현
- 데이터에 없는 내용 추측 (학교생활, 수업 태도 등)"""


def format_student_for_prompt(student: dict) -> str:
    """학생 데이터를 Claude가 읽기 쉬운 텍스트로 변환."""
    lines = []
    lines.append(f"[학생 정보]")
    lines.append(f"학년/반: {student['code']} ({student['school_level']} {student['grade']}학년)")
    lines.append(f"이름: {student['name']}")

    # 진도
    if student.get("progress"):
        lines.append(f"\n[이번 달 진도 계획]")
        for r in student["progress"]["roles"]:
            교재 = r.get("교재") or "교재 미정"
            lines.append(f"- {r['role']}: {교재}")
            주차별 = {int(k): v for k, v in r.get("주차별", {}).items()}
            if 주차별:
                단원들 = list(dict.fromkeys(주차별.values()))
                lines.append(f"  진도: {' → '.join(단원들)}")

    # 시험대비 플래그
    tp_plan = student.get("test_prep_plan")
    tp_done = student.get("test_prep_done")
    if tp_plan and tp_done:
        lines.append(f"\n[※※※ 시험대비 완료 + 예정 — 과거형과 미래형 혼합 사용 ※※※]")
    elif tp_plan:
        lines.append(f"\n[※※※ 시험대비 예정 — 미래형으로 '~할 계획입니다' 사용 ※※※]")
    elif tp_done:
        lines.append(f"\n[※※※ 시험대비 완료 — 과거형으로 '~했습니다' 사용 ※※※]")

    # 재시 예정 학습지
    pending = student.get("pending_retests") or []
    if pending:
        lines.append(f"\n[재시 예정 (70점 미만이고 아직 재시 못 친 학습지)]")
        for p in pending:
            주차 = f"{p['주차']}주차 " if p.get("주차") else ""
            lines.append(f"- {주차}{p['단원명']}: {p['점수']}점 → 오는 달에 재시 예정")

    # 학습 분석 (난이도, 학습량, 강점/약점)
    analytics = student.get("analytics") or {}
    if analytics.get("total_count"):
        lines.append(f"\n[이번 달 학습량]")
        lines.append(f"- 교재 외 학습지: {analytics['total_count']}건 ({analytics.get('total_problems', 0)}문항)")

        by_diff = analytics.get("by_difficulty") or {}
        if by_diff:
            lines.append(f"\n[난이도별 학습 분포]")
            for d in sorted(by_diff.keys(), key=lambda x: int(x)):
                info = by_diff[d]
                lines.append(
                    f"- 난이도 {d} ({info['label']}): {info['count']}건, "
                    f"평균 {info['avg_score']}점, {info['total_problems']}문항"
                )

        if analytics.get("strengths"):
            lines.append(f"\n[잘한 단원 (85점 이상)]")
            for s in analytics["strengths"]:
                lines.append(f"- {s['단원']}: 평균 {s['avg_score']}점 ({s['count']}회)")

        if analytics.get("weaknesses"):
            lines.append(f"\n[부족한 단원 (70점 미만)]")
            for w in analytics["weaknesses"]:
                lines.append(f"- {w['단원']}: 평균 {w['avg_score']}점 ({w['count']}회)")

    # 주간평가 (가통문에 표시되는 것)
    if student.get("weekly_evals"):
        lines.append(f"\n[주간 평가 결과 (가통문 표시용)]")
        for e in student["weekly_evals"]:
            재시 = f" (재시{e['재시차수']}차)" if e.get("재시차수", 0) > 0 else ""
            주차 = f"{e.get('주차', '?')}주차"
            tag = " ← 70점 미만" if e["점수"] is not None and e["점수"] < 70 else ""
            lines.append(f"- {주차} {e['단원명']}{재시}: {e['점수']}점{tag}")

    # 선행평가
    if student.get("unit_evals"):
        lines.append(f"\n[선행 평가 결과 (가통문 표시용)]")
        for e in student["unit_evals"]:
            재시 = f" (재시{e['재시차수']}차)" if e.get("재시차수", 0) > 0 else ""
            tag = " ← 70점 미만" if e["점수"] is not None and e["점수"] < 70 else ""
            lines.append(f"- {e['단원명']}{재시}: {e['점수']}점{tag}")

    return "\n".join(lines)


def generate_comment(client: Anthropic, student: dict) -> str:
    """한 학생의 코멘트 생성."""
    user_msg = format_student_for_prompt(student)

    resp = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=[
            {
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[{"role": "user", "content": user_msg}],
    )
    return resp.content[0].text.strip(), resp.usage


def main():
    if len(sys.argv) < 3:
        print("사용법: python generate_comments.py <연도> <월>")
        sys.exit(1)

    year = int(sys.argv[1])
    month = int(sys.argv[2])
    force = "--force" in sys.argv  # 기존 코멘트 있어도 재생성

    if not os.getenv("ANTHROPIC_API_KEY"):
        print("❌ ANTHROPIC_API_KEY 없음. .env 확인.")
        sys.exit(1)

    merged_path = BASE / "output" / f"merged_{year}_{month:02d}.json"
    if not merged_path.exists():
        print(f"❌ {merged_path.name} 없음")
        sys.exit(1)

    with merged_path.open(encoding="utf-8") as f:
        data = json.load(f)

    client = Anthropic()

    total_input = 0
    total_output = 0
    total_cache_read = 0
    total_cache_create = 0
    generated = 0
    skipped = 0

    for i, student in enumerate(data["students"], 1):
        # 데이터 없는 학생 skip
        if not (student.get("weekly_evals") or student.get("unit_evals") or student.get("progress")):
            continue

        if student.get("ai_comment") and not force:
            skipped += 1
            continue

        label = f"[{i}/{len(data['students'])}] {student['code']} {student['name']}"
        print(f"  생성 중: {label}", flush=True)

        try:
            comment, usage = generate_comment(client, student)
            student["ai_comment"] = comment
            generated += 1

            total_input += getattr(usage, "input_tokens", 0)
            total_output += getattr(usage, "output_tokens", 0)
            total_cache_read += getattr(usage, "cache_read_input_tokens", 0) or 0
            total_cache_create += getattr(usage, "cache_creation_input_tokens", 0) or 0
        except Exception as e:
            print(f"    ⚠️ 실패: {e}")
            continue

        # 중간 저장 (장시간 실행 대비)
        if generated % 10 == 0:
            with merged_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        time.sleep(0.3)  # API rate limit 여유

    # 최종 저장
    with merged_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 비용 예상 (Sonnet 4.5: $3/M input, $15/M output, cached read $0.30/M)
    cost_input = total_input * 3 / 1_000_000
    cost_output = total_output * 15 / 1_000_000
    cost_cache_create = total_cache_create * 3.75 / 1_000_000
    cost_cache_read = total_cache_read * 0.30 / 1_000_000
    total_cost = cost_input + cost_output + cost_cache_create + cost_cache_read

    print(f"\n✅ 완료")
    print(f"   생성: {generated}명 / 스킵(이미 있음): {skipped}명")
    print(f"   토큰: input={total_input:,} / output={total_output:,} / "
          f"cache_read={total_cache_read:,} / cache_create={total_cache_create:,}")
    print(f"   예상 비용: ${total_cost:.4f} (약 {total_cost*1350:.0f}원)")
    print(f"   캐시 효율: {total_cache_read/(total_input+total_cache_read)*100:.0f}% 캐시 재사용"
          if (total_input + total_cache_read) > 0 else "")

    # 샘플 출력
    print(f"\n   📝 샘플 코멘트:")
    for s in data["students"][:3]:
        if s.get("ai_comment"):
            print(f"\n   [{s['code']} {s['name']}]")
            print(f"   {s['ai_comment']}")


if __name__ == "__main__":
    main()
