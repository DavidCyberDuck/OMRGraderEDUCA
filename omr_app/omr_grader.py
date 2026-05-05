"""
OMR Grader — scores MC answers; passes through SK, words, Grado, Grupo, Folio.
"""
from dataclasses import dataclass, field
from typing import Optional
from omr_scanner import ScanResult


@dataclass
class GradeResult:
    page_num:        int
    folio:           str             # scanned 2-digit folio string e.g. '37'
    grado:           Optional[str]
    grupo:           Optional[str]
    mc_answers:      list
    mc_correct:      list
    score:           int
    total:           int
    percentage:      float
    sk_answers:      list
    sk_average:      Optional[float]
    word_selections: list            # list[bool], one per SEC3_WORDS entry
    confidence:      float
    error:           Optional[str] = None


def grade_results(scan_results, answer_key):
    graded = []
    for scan in scan_results:
        words = getattr(scan, 'word_selections', [])
        if scan.error:
            graded.append(GradeResult(
                page_num=scan.page_num, folio=scan.folio,
                grado=scan.grado, grupo=scan.grupo,
                mc_answers=scan.mc_answers, mc_correct=[],
                score=0, total=len(answer_key), percentage=0.0,
                sk_answers=scan.sk_answers, sk_average=None,
                word_selections=words,
                confidence=scan.confidence, error=scan.error,
            ))
            continue

        mc_correct, score = [], 0
        for i, ans in enumerate(scan.mc_answers):
            if i < len(answer_key):
                ok = (ans == answer_key[i]) if ans is not None else False
                mc_correct.append(ok)
                if ok:
                    score += 1
            else:
                mc_correct.append(False)

        total  = len(answer_key)
        pct    = round(score / total * 100, 1) if total > 0 else 0.0
        sk     = [v for v in scan.sk_answers if v is not None]
        sk_avg = round(sum(sk)/len(sk), 2) if sk else None

        graded.append(GradeResult(
            page_num=scan.page_num, folio=scan.folio,
            grado=scan.grado, grupo=scan.grupo,
            mc_answers=scan.mc_answers, mc_correct=mc_correct,
            score=score, total=total, percentage=pct,
            sk_answers=scan.sk_answers, sk_average=sk_avg,
            word_selections=words,
            confidence=scan.confidence,
        ))
    return graded
