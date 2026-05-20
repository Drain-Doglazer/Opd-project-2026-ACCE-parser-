import json
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from run_parser import run_pipeline

# ── Config ───────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="ACCE Equipment List Parser",
    page_icon="⚙️",
    layout="centered",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": """
## Требования к файлам

**Перечень оборудования** — поддерживаемые форматы:
- `.xlsx` — таблица Excel с колонками тегов, наименований и параметров оборудования
- `.pdf` — документ с таблицей перечня оборудования
- `.docx` — документ Word с таблицей перечня оборудования

**Шаблон ACCE** (опционально):
- `.xlsx` — master-template, экспортированный из Aspen Capital Cost Estimator

**Ограничения:**
- Файл должен содержать хотя бы одну таблицу с колонкой тегов оборудования
- Кодировка текста: UTF-8 или Windows-1251
""",
    },
)

st.markdown(
    """
    <style>
    button[title="Print"] {display: none !important;}
    button[title="Record a screencast"] {display: none !important;}
    button[title="Clear cache"] {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ── History helpers ───────────────────────────────────────────────────────────

_HISTORY_FILE = Path(__file__).parent / "data" / "history.json"


def _load_history() -> list[dict]:
    if _HISTORY_FILE.exists():
        try:
            return json.loads(_HISTORY_FILE.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _append_history(entry: dict) -> None:
    history = _load_history()
    history.insert(0, entry)
    _HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    _HISTORY_FILE.write_text(
        json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── App ───────────────────────────────────────────────────────────────────────

st.title("⚙️ ACCE Equipment List Parser")
st.caption("Автоматический парсер перечней оборудования для Aspen Capital Cost Estimator")

tab_parser, tab_history = st.tabs(["Парсер", "История"])

# ════════════════════════════════════════════════════════════════════════════
# TAB: Парсер
# ════════════════════════════════════════════════════════════════════════════
with tab_parser:

    # ── Inputs ───────────────────────────────────────────────────────────────
    st.subheader("1. Загрузите файлы")

    input_file = st.file_uploader(
        "Перечень оборудования",
        type=["xlsx", "pdf", "docx"],
        help="Поддерживаются форматы Excel, PDF и Word",
    )

    template_file = st.file_uploader(
        "Шаблон ACCE (master-template.xlsx)",
        type=["xlsx"],
        help="Опционально. Если загружен — создаётся файл для импорта в ACCE",
    )

    area_name = st.text_input(
        "Зона установки (Parent Area)",
        value="Main Area",
        help="Используется как PARENT_AREA для записей без указания зоны",
    )

    # ── Run ──────────────────────────────────────────────────────────────────
    st.subheader("2. Запустите парсер")

    run_btn = st.button("▶ Запустить", type="primary", use_container_width=True)

    if run_btn and input_file:
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = str(Path(tmpdir) / input_file.name)
            with open(input_path, "wb") as f:
                f.write(input_file.getvalue())

            template_path = None
            if template_file:
                template_path = str(Path(tmpdir) / template_file.name)
                with open(template_path, "wb") as f:
                    f.write(template_file.getvalue())

            parsed_path = str(Path(tmpdir) / "parsed.xlsx")
            output_path = str(Path(tmpdir) / "acce_ready.xlsx")

            with st.spinner("Обработка файла..."):
                try:
                    result = run_pipeline(
                        input_path=input_path,
                        template_path=template_path,
                        area=area_name,
                        output_path=output_path,
                        parsed_path=parsed_path,
                    )
                except Exception as e:
                    st.error(f"Ошибка: {e}")
                    st.stop()

            # ── Save to history ───────────────────────────────────────────
            hist_entry: dict = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "filename":  input_file.name,
                "total":     result["total"],
                "valid":     result["valid"],
                "invalid":   result["invalid"],
                "warnings":  result["warnings"],
                "with_template": bool(template_file),
            }
            summary = result.get("summary")
            if summary:
                hist_entry["exported"] = summary.exported
                hist_entry["skipped"]  = summary.skipped_invalid
                hist_entry["unknown"]  = summary.unknown_symbol
            _append_history(hist_entry)

            # ── Results ──────────────────────────────────────────────────
            st.subheader("3. Результат")

            df_parsed = pd.read_excel(parsed_path, dtype=str).fillna("")

            res_stats, res_preview, res_assess = st.tabs(
                ["Сводка", "Предпросмотр", "Оценка"]
            )

            with res_stats:
                col1, col2, col3 = st.columns(3)
                col1.metric("Всего записей", result["total"])
                col2.metric("Валидных",      result["valid"])
                col3.metric(
                    "Невалидных",
                    result["invalid"],
                    delta=f"-{result['invalid']}" if result["invalid"] else None,
                    delta_color="inverse",
                )

                if result["warnings"]:
                    st.info(f"⚠️ Записей с предупреждениями: {result['warnings']}")

                if summary:
                    st.divider()
                    st.markdown("**Экспорт в ACCE**")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Экспортировано", summary.exported)
                    c2.metric("Пропущено",      summary.skipped_invalid)
                    c3.metric("UNKNOWN",        summary.unknown_symbol)

                    if summary.by_sheet:
                        with st.expander("По листам ACCE"):
                            for sheet, count in sorted(
                                summary.by_sheet.items(), key=lambda x: -x[1]
                            ):
                                st.write(f"**{sheet}**: {count} записей")

            with res_preview:
                _PREVIEW_COLS = [
                    "USER_TAG", "DESCRIPTION_RU", "PARENT_AREA",
                    "ACCE_ITEM_SYMBOL", "ACCE_ITEM_TYPE", "QUANTITY",
                    "IS_VALID", "ERRORS", "WARNINGS",
                ]
                df_show = df_parsed[
                    [c for c in _PREVIEW_COLS if c in df_parsed.columns]
                ].copy()

                status_filter = st.radio(
                    "Показать",
                    ["Все", "Валидные", "Невалидные", "С предупреждениями"],
                    horizontal=True,
                )
                if status_filter == "Валидные":
                    df_show = df_show[df_show["IS_VALID"] == "YES"]
                elif status_filter == "Невалидные":
                    df_show = df_show[df_show["IS_VALID"] == "NO"]
                elif status_filter == "С предупреждениями":
                    df_show = df_show[df_show["WARNINGS"].str.strip() != ""]

                def _color_row(row):
                    if row.get("IS_VALID") == "NO":
                        bg = "#ffd6d6"
                    elif str(row.get("WARNINGS", "")).strip():
                        bg = "#fff3cd"
                    else:
                        bg = "#d4edda"
                    return [f"background-color: {bg}"] * len(row)

                st.dataframe(
                    df_show.style.apply(_color_row, axis=1),
                    use_container_width=True,
                    height=420,
                )

            with res_assess:
                total = result["total"]
                if total == 0:
                    st.warning("Нет записей для оценки.")
                else:
                    valid_pct   = result["valid"]    / total
                    invalid_pct = result["invalid"]  / total
                    warn_pct    = result["warnings"] / total

                    # ── Overall score ─────────────────────────────────
                    score = round(valid_pct * 100, 1)
                    grade = (
                        "Отлично"           if score >= 90 else
                        "Хорошо"            if score >= 70 else
                        "Требует доработки"
                    )

                    sc1, sc2, sc3, sc4 = st.columns(4)
                    sc1.metric("Качество парсинга", f"{score}%", grade)
                    sc2.metric("Всего записей",     total)
                    sc3.metric("Валидных",          result["valid"])
                    sc4.metric("Невалидных",        result["invalid"])

                    st.divider()

                    # ── Validity progress bars ────────────────────────
                    st.markdown("**Валидность записей**")

                    _validity_rows = [
                        ("Валидные",           result["valid"],    valid_pct,   "#d4edda"),
                        ("Невалидные",         result["invalid"],  invalid_pct, "#ffd6d6"),
                        ("С предупреждениями", result["warnings"], warn_pct,    "#fff3cd"),
                    ]
                    for label, count, pct, _ in _validity_rows:
                        lc, bc, nc = st.columns([2, 6, 1])
                        lc.write(f"{label} ({count})")
                        bc.progress(float(pct))
                        nc.write(f"**{pct*100:.1f}%**")

                    # ── Classification breakdown ──────────────────────
                    if "ACCE_ITEM_SYMBOL" in df_parsed.columns:
                        st.divider()
                        st.markdown("**Классификация ACCE**")

                        sym = df_parsed["ACCE_ITEM_SYMBOL"].str.strip().str.upper()
                        unknown_mask = sym.isin(["", "UNKNOWN"])
                        n_unknown  = int(unknown_mask.sum())
                        n_known    = total - n_unknown
                        known_pct   = n_known   / total
                        unknown_pct = n_unknown / total

                        _cls_rows = [
                            ("Распознано",    n_known,   known_pct),
                            ("Не распознано", n_unknown, unknown_pct),
                        ]
                        for label, count, pct in _cls_rows:
                            lc, bc, nc = st.columns([2, 6, 1])
                            lc.write(f"{label} ({count})")
                            bc.progress(float(pct))
                            nc.write(f"**{pct*100:.1f}%**")

                        if n_known > 0:
                            with st.expander("По символам ACCE"):
                                sym_counts = (
                                    df_parsed[~unknown_mask]["ACCE_ITEM_SYMBOL"]
                                    .value_counts()
                                    .rename_axis("Символ")
                                    .reset_index(name="Кол-во")
                                )
                                sym_counts["% от всех"] = (
                                    sym_counts["Кол-во"] / total * 100
                                ).round(1).astype(str) + "%"
                                st.dataframe(
                                    sym_counts,
                                    use_container_width=True,
                                    hide_index=True,
                                )

                        if n_unknown > 0:
                            with st.expander("Нераспознанные записи"):
                                unk_cols = [
                                    c for c in
                                    ["USER_TAG", "DESCRIPTION_RU", "ACCE_ITEM_TYPE", "ERRORS"]
                                    if c in df_parsed.columns
                                ]
                                st.dataframe(
                                    df_parsed[unknown_mask][unk_cols],
                                    use_container_width=True,
                                    hide_index=True,
                                )

            # ── Downloads ────────────────────────────────────────────────
            st.divider()
            st.subheader("4. Скачайте результат")

            dl1, dl2 = st.columns(2)

            with open(parsed_path, "rb") as f:
                dl1.download_button(
                    label="⬇ Скачать parsed.xlsx",
                    data=f.read(),
                    file_name=f"{Path(input_file.name).stem}_parsed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            if result["output_path"] and Path(output_path).exists():
                with open(output_path, "rb") as f:
                    dl2.download_button(
                        label="⬇ Скачать acce_ready.xlsx",
                        data=f.read(),
                        file_name=f"{Path(input_file.name).stem}_acce.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

    elif run_btn and not input_file:
        st.warning("Загрузите файл перечня оборудования")

# ════════════════════════════════════════════════════════════════════════════
# TAB: История
# ════════════════════════════════════════════════════════════════════════════
with tab_history:
    history = _load_history()

    if not history:
        st.info("История пустая. Запустите парсер — здесь появится запись о каждом файле.")
    else:
        h_top, h_clear = st.columns([5, 1])
        h_top.markdown(f"Всего запусков: **{len(history)}**")
        if h_clear.button("Очистить", type="secondary", use_container_width=True):
            _HISTORY_FILE.unlink(missing_ok=True)
            st.rerun()

        st.divider()

        _HIST_RENAME = {
            "timestamp":     "Дата/время",
            "filename":      "Файл",
            "total":         "Всего",
            "valid":         "Валидных",
            "invalid":       "Невалидных",
            "warnings":      "Предупр.",
            "with_template": "Шаблон",
            "exported":      "Экспорт",
            "skipped":       "Пропущено",
            "unknown":       "UNKNOWN",
        }

        df_hist = (
            pd.DataFrame(history)
            .rename(columns=_HIST_RENAME)
        )
        if "Шаблон" in df_hist.columns:
            df_hist["Шаблон"] = df_hist["Шаблон"].map({True: "да", False: "нет"})

        def _color_hist_row(row):
            inv = row.get("Невалидных", 0)
            warn = row.get("Предупр.", 0)
            try:
                if int(inv) > 0:
                    bg = "#ffd6d6"
                elif int(warn) > 0:
                    bg = "#fff3cd"
                else:
                    bg = "#d4edda"
            except (ValueError, TypeError):
                bg = "#ffffff"
            return [f"background-color: {bg}"] * len(row)

        st.dataframe(
            df_hist.style.apply(_color_hist_row, axis=1),
            use_container_width=True,
            hide_index=True,
        )
