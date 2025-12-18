from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import List

import pandas as pd

from config import AppConfig
from core import db

log = logging.getLogger("naos_sla.recommend")


try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
except Exception:  # pragma: no cover - optional dependency handled gracefully
    TfidfVectorizer = None  # type: ignore
    cosine_similarity = None  # type: ignore


@dataclass
class RecommendationRule:
    question: str
    answer: str


def _safe_vectorizer(max_features: int = 4000) -> TfidfVectorizer | None:
    if TfidfVectorizer is None:
        log.warning("sklearn not available, recommendations disabled")
        return None
    return TfidfVectorizer(
        max_features=max_features, ngram_range=(1, 2), stop_words=None
    )


def load_qa_pairs(path: Path) -> List[RecommendationRule]:
    if not path or not path.exists():
        return []
    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception as exc:
        log.warning("Failed to load ORM script: %s", exc)
        return []
    cols = {c.lower(): c for c in df.columns}
    q_col = None
    a_col = None
    for key, col in cols.items():
        if ("вопрос" in key) or ("question" in key):
            q_col = col
        if ("ответ" in key) or ("answer" in key) or ("скрипт" in key):
            a_col = col
    if not q_col or not a_col:
        log.warning("ORM script does not have question/answer columns")
        return []
    rules: List[RecommendationRule] = []
    for _, row in df.iterrows():
        q = str(row.get(q_col, "") or "").strip()
        a = str(row.get(a_col, "") or "").strip()
        if q and a:
            rules.append(RecommendationRule(question=q, answer=a))
    return rules


def update_recommendations(cfg: AppConfig) -> None:
    """Compute repeat hints and recommended answers (offline TF-IDF)."""
    vec = _safe_vectorizer()
    if vec is None or cosine_similarity is None:
        return
    lookback = cfg.orm_similarity_days
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    rows = conn.execute(
        """
        SELECT id, subject, body FROM tickets
        WHERE datetime(first_received_utc) >= datetime('now', ?)
        ORDER BY datetime(first_received_utc) DESC
        """,
        (f"-{lookback} day",),
    ).fetchall()
    if not rows:
        conn.close()
        return
    texts = [f"{r['subject']} {r['body']}" for r in rows]
    try:
        matrix = vec.fit_transform(texts)
    except Exception as exc:
        log.warning("TF-IDF build failed: %s", exc)
        conn.close()
        return

    qa_rules = load_qa_pairs(
        Path(cfg.orm_script_path) if cfg.orm_script_path else cfg.paths.orm_script_path
    )
    qa_vec = _safe_vectorizer()
    qa_matrix = None
    if qa_vec and qa_rules:
        try:
            qa_matrix = qa_vec.fit_transform([r.question for r in qa_rules])
        except Exception as exc:
            log.warning("TF-IDF for QA failed: %s", exc)
            qa_matrix = None

    for idx, row in enumerate(rows):
        sims = cosine_similarity(matrix[idx], matrix).flatten()
        sims[idx] = 0.0
        matches = [
            (score, rows[i])
            for i, score in enumerate(sims)
            if score >= cfg.orm_similarity_threshold
        ]
        matches.sort(key=lambda x: x[0], reverse=True)
        repeat_hint = ""
        is_repeat = False
        if matches:
            freq = len(matches)
            top_subj = matches[0][1]["subject"]
            repeat_hint = f"{freq} похожих за {lookback}д (топ: {top_subj})"
            is_repeat = True

        recommended_answer = ""
        topic = None
        match_score = None
        if qa_matrix is not None and qa_vec is not None and qa_rules:
            try:
                ticket_vec = qa_vec.transform([texts[idx]])
                qa_scores = cosine_similarity(ticket_vec, qa_matrix).flatten()
                best_pairs = sorted(
                    enumerate(qa_scores), key=lambda x: x[1], reverse=True
                )[: cfg.orm_max_suggestions]
                best_pairs = [
                    p for p in best_pairs if p[1] >= cfg.orm_similarity_threshold
                ]
                if best_pairs:
                    best_idx, best_score = best_pairs[0]
                    recommended_answer = qa_rules[best_idx].answer
                    topic = qa_rules[best_idx].question
                    match_score = float(best_score)
            except Exception as exc:
                log.debug("QA recommendation failed: %s", exc)

        conn.execute(
            """
            UPDATE tickets
            SET is_repeat=?, repeat_hint=?, recommended_answer=?, match_score=?, topic=?
            WHERE id=?
            """,
            (
                int(is_repeat),
                repeat_hint,
                recommended_answer,
                match_score,
                topic,
                row["id"],
            ),
        )
    conn.commit()
    conn.close()
    log.info("Recommendations refreshed for %s tickets", len(rows))
