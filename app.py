"""
ОКПД 2 — Сопоставление кодов с ПП 1875
Базы данных лежат РЯДОМ с этим файлом (автозагрузка, без ручного выбора).
Запуск: streamlit run app.py
"""

import re, io
from pathlib import Path
import streamlit as st
import pandas as pd

# ─────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────
st.set_page_config(page_title="ОКПД 2 · ПП 1875", page_icon="🗂️", layout="wide")

BASE_DIR  = Path(__file__).parent          # папка рядом с app.py
FILE_MAIN = BASE_DIR / "Код_ОКПД_2.xlsx"
FILE_APP1 = BASE_DIR / "Приложение_N_1.xlsx"
FILE_APP2 = BASE_DIR / "Приложение_N_2.xlsx"

# ─────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Manrope:wght@400;600;700;800&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #0F172A !important; color: #CBD5E1;
    font-family: 'Manrope', sans-serif;
}
[data-testid="stSidebar"]  { background: #090E1A !important; }
[data-testid="stHeader"]   { background: transparent !important; }
h1,h2,h3 { color: #F1F5F9 !important; font-weight: 800; }

[data-baseweb="tab-list"]  { background: #1E293B !important; border-radius: 12px; padding: 4px; gap: 4px; }
[data-baseweb="tab"]       { border-radius: 8px !important; color: #94A3B8 !important; font-weight: 600 !important; }
[aria-selected="true"]     { background: #334155 !important; color: #F1F5F9 !important; }

input, [data-baseweb="input"] input {
    background: #1E293B !important; border: 1px solid #334155 !important;
    color: #F1F5F9 !important; border-radius: 8px !important;
    font-family: 'JetBrains Mono', monospace !important;
}
[data-testid="stFileUploader"] {
    background: #1E293B !important; border: 2px dashed #334155 !important;
    border-radius: 12px !important; padding: 16px !important;
}
[data-testid="baseButton-primary"] {
    background: linear-gradient(135deg,#2563EB,#1D4ED8) !important;
    border: none !important; color: #fff !important; font-weight: 700 !important;
    border-radius: 8px !important; box-shadow: 0 4px 15px rgba(37,99,235,.35) !important;
    transition: all .2s !important;
}
[data-testid="baseButton-primary"]:hover { transform: translateY(-1px); }
[data-testid="baseButton-secondary"] {
    background: transparent !important; border: 1px solid #334155 !important;
    color: #64748B !important; border-radius: 8px !important; transition: all .2s;
}
[data-testid="baseButton-secondary"]:hover {
    background: rgba(239,68,68,.15) !important; border-color: #EF4444 !important; color: #EF4444 !important;
}
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg,#059669,#047857) !important;
    border: none !important; color: #fff !important; font-weight: 700 !important;
    border-radius: 8px !important; box-shadow: 0 4px 15px rgba(5,150,105,.35) !important;
}
[data-testid="stDownloadButton"] button:hover { transform: translateY(-1px); }

[data-testid="stMetric"]      { background:#1E293B !important; border:1px solid #334155 !important; border-radius:12px !important; padding:18px !important; }
[data-testid="stMetricValue"] { color:#38BDF8 !important; font-weight:800 !important; }
[data-testid="stMetricLabel"] { color:#94A3B8 !important; }
[data-testid="stDataFrame"]   { border-radius:12px; overflow:hidden; }

.card {
    background:#1E293B; border:1px solid #334155; border-left:4px solid #2563EB;
    border-radius:12px; padding:16px 20px; margin:8px 0;
}
.card.yes { border-left-color:#10B981; }
.card.no  { border-left-color:#EF4444; }
.card-lbl { color:#64748B; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.08em; margin-bottom:4px; }
.card-val { color:#F1F5F9; font-size:17px; font-weight:800; margin-bottom:6px; }
.card-sub { color:#94A3B8; font-size:13px; font-family:'JetBrains Mono',monospace; }
.badge    { display:inline-block; padding:2px 10px; border-radius:20px; font-size:12px; font-weight:700; margin-left:8px; }
.b-yes    { background:rgba(16,185,129,.15); color:#10B981; border:1px solid #10B981; }
.b-no     { background:rgba(239,68,68,.12);  color:#EF4444; border:1px solid #EF4444; }
.mono     { font-family:'JetBrains Mono',monospace; color:#38BDF8; background:#0F172A; padding:2px 8px; border-radius:4px; font-size:13px; }
.ok-pill  { background:rgba(16,185,129,.12); color:#10B981; border:1px solid rgba(16,185,129,.3); border-radius:8px; padding:8px 14px; font-size:14px; font-weight:600; display:inline-block; margin-bottom:8px; }
hr { border-color:#1E293B !important; }
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:#0F172A; }
::-webkit-scrollbar-thumb { background:#334155; border-radius:3px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
# УТИЛИТЫ
# ─────────────────────────────────────────────────────────
def clean_code(raw: object) -> str:
    return re.sub(r'[^\d.]', '', str(raw)).strip('.')


@st.cache_data(show_spinner="⏳ Загружаю справочник ОКПД 2…")
def load_main() -> dict:
    """Статус = код ОКПД 2, Название = наименование."""
    df = pd.read_excel(FILE_MAIN, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    code_col = 'Статус'
    name_col = 'Название'
    result = {}
    for _, row in df.iterrows():
        c = clean_code(row[code_col])
        n = str(row[name_col]).strip()
        if c and n.lower() not in ('nan', 'none', ''):
            result[c] = n
    return result


@st.cache_data(show_spinner="⏳ Загружаю Приложения…")
def load_appendix(path: str) -> dict:
    """skiprows=1 — пропускаем строку-заголовок; col[2]=код, col[1]=наим."""
    df = pd.read_excel(path, header=None, dtype=str, skiprows=1)
    result = {}
    for _, row in df.iterrows():
        if len(row) < 3:
            continue
        c = clean_code(row.iloc[2])
        n = str(row.iloc[1]).strip()
        if c and n.lower() not in ('nan', 'none', ''):
            result[c] = n
    return result


def hier_lookup(code: str, lookup: dict) -> tuple[bool, str]:
    """
    1. Точное совпадение.
    2. code начинается на ключ словаря (длинный код → ищем корень).
    3. Ключ словаря начинается на code (короткий код → группа).
    """
    if not code:
        return False, ''
    if code in lookup:
        return True, lookup[code]
    # ищем самый длинный совпадающий корень
    best, best_len = '', 0
    for key in lookup:
        if code.startswith(key) and len(key) > best_len:
            best, best_len = key, len(key)
    if best:
        return True, lookup[best]
    # code — группа, берём первый попавшийся дочерний
    for key in lookup:
        if key.startswith(code):
            return True, lookup[key]
    return False, ''


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────
for k in ('hist_code', 'hist_name'):
    if k not in st.session_state:
        st.session_state[k] = []


# ─────────────────────────────────────────────────────────
# ПРОВЕРКА НАЛИЧИЯ ФАЙЛОВ
# ─────────────────────────────────────────────────────────
missing_files = [f.name for f in (FILE_MAIN, FILE_APP1, FILE_APP2) if not f.exists()]
if missing_files:
    st.error(f"❌ Файлы не найдены рядом с app.py: {', '.join(missing_files)}\n\n"
             "Положите их в одну папку с app.py и перезапустите.")
    st.stop()

# Загрузка (кэшируется)
main_d = load_main()
app1_d = load_appendix(str(FILE_APP1))
app2_d = load_appendix(str(FILE_APP2))


# ─────────────────────────────────────────────────────────
# ЗАГОЛОВОК
# ─────────────────────────────────────────────────────────
c_ttl, c_stat = st.columns([3, 1])
with c_ttl:
    st.markdown("""
    <h1 style="margin-bottom:2px">🗂️ ОКПД 2 · Постановление № 1875</h1>
    <p style="color:#64748B;margin-top:0;font-size:15px">
    Автоматическое сопоставление кодов со справочником и Приложениями №1, №2
    </p>""", unsafe_allow_html=True)
with c_stat:
    st.markdown(f"""
    <div style="background:#1E293B;border:1px solid #334155;border-radius:12px;padding:14px 16px;margin-top:8px">
        <div style="color:#64748B;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em">Базы загружены</div>
        <div style="color:#38BDF8;font-size:15px;font-weight:800;margin-top:4px">
            ОКПД 2: {len(main_d):,} кодов
        </div>
        <div style="color:#94A3B8;font-size:12px;margin-top:2px">
            Прил.№1: {len(app1_d)} поз. &nbsp;·&nbsp; Прил.№2: {len(app2_d)} поз.
        </div>
    </div>""", unsafe_allow_html=True)

st.markdown("---")


# ─────────────────────────────────────────────────────────
# ВКЛАДКИ
# ─────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📂 Массовая проверка", "🔎 Поиск по коду", "🔤 Поиск по названию"])


# ══════════════════════════════════════════════════════════
# TAB 1 — МАССОВАЯ ПРОВЕРКА
# ══════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Загрузите рабочий файл")
    st.caption(
        "Файл должен содержать колонку **Код ОКПД 2**. "
        "Все исходные колонки сохраняются. В конец добавляются три новых."
    )

    user_file = st.file_uploader(
        "Рабочий файл (.xlsx / .xls)", type=["xlsx", "xls"], key="user_f"
    )

    if user_file:
        user_df = pd.read_excel(user_file, dtype=str)
        user_df.columns = [str(c).strip() for c in user_df.columns]

        # Найти колонку с кодом ОКПД 2
        okpd_col = next(
            (c for c in user_df.columns if re.search(r'окпд.*2|код.*окпд', c.lower())),
            next((c for c in user_df.columns if 'окпд' in c.lower()), None)
        )

        if okpd_col is None:
            st.error(f"❌ Колонка «Код ОКПД 2» не найдена. Доступные колонки: {list(user_df.columns)}")
        else:
            col_info, col_btn = st.columns([4, 1])
            col_info.info(f"Колонка кодов: **{okpd_col}** · Строк: **{len(user_df):,}**")

            if col_btn.button("🚀 Сопоставить", type="primary"):
                names_res, app1_res, app2_res = [], [], []
                bar   = st.progress(0, text="Обрабатываю строки…")
                total = len(user_df)

                for i, raw in enumerate(user_df[okpd_col]):
                    code = clean_code(raw)

                    # 1. Наименование — ТОЛЬКО из главного справочника
                    fm, vm = hier_lookup(code, main_d)
                    names_res.append(vm if fm else '')

                    # 2. Приложение №1 — ДА/НЕТ
                    f1, _ = hier_lookup(code, app1_d)
                    app1_res.append('ДА' if f1 else 'НЕТ')

                    # 3. Приложение №2 — ДА/НЕТ
                    f2, _ = hier_lookup(code, app2_d)
                    app2_res.append('ДА' if f2 else 'НЕТ')

                    if i % 100 == 0:
                        bar.progress(min(i / total, 1.0), text=f"Обработано {i:,} / {total:,}…")

                bar.empty()

                # Добавляем три колонки в конец — исходные не трогаем
                result = user_df.copy()
                result['Наименование (ОКПД 2)']   = names_res
                result['В Приложении №1 (ПП 1875)'] = app1_res
                result['В Приложении №2 (ПП 1875)'] = app2_res

                # ── Статистика ──────────────────────────────────────
                found_m  = sum(1 for v in names_res if v)
                cnt_a1   = app1_res.count('ДА')
                cnt_a2   = app2_res.count('ДА')
                cnt_both = sum(1 for a, b in zip(app1_res, app2_res) if a == 'ДА' and b == 'ДА')
                pct      = lambda n: f"{n/total*100:.0f}%"

                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("📋 Всего строк",          f"{total:,}")
                c2.metric("📚 Найдено в ОКПД 2",     f"{found_m:,}", pct(found_m))
                c3.metric("✅ В Приложении №1",      f"{cnt_a1:,}",  pct(cnt_a1))
                c4.metric("✅ В Приложении №2",      f"{cnt_a2:,}",  pct(cnt_a2))
                c5.metric("🔗 В обоих Приложениях",  f"{cnt_both:,}", pct(cnt_both))

                # ── Предпросмотр ─────────────────────────────────────
                st.markdown("#### Предпросмотр (первые 10 строк)")
                preview_cols = (
                    list(user_df.columns)
                    + ['Наименование (ОКПД 2)', 'В Приложении №1 (ПП 1875)', 'В Приложении №2 (ПП 1875)']
                )
                # Подсветка строк с ДА
                def color_rows(row):
                    a1 = row.get('В Приложении №1 (ПП 1875)', '')
                    a2 = row.get('В Приложении №2 (ПП 1875)', '')
                    if a1 == 'ДА' and a2 == 'ДА':
                        return ['background-color: rgba(251,191,36,.12)'] * len(row)
                    if a1 == 'ДА':
                        return ['background-color: rgba(16,185,129,.08)'] * len(row)
                    if a2 == 'ДА':
                        return ['background-color: rgba(56,189,248,.08)'] * len(row)
                    return [''] * len(row)

                st.dataframe(
                    result[preview_cols].head(10).style.apply(color_rows, axis=1),
                    use_container_width=True,
                )

                # ── Скачать ──────────────────────────────────────────
                st.download_button(
                    "⬇️ Скачать полный результат (.xlsx)",
                    data=to_excel(result[preview_cols]),
                    file_name="okpd2_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    else:
        st.markdown("""
        <div class="card">
            <div class="card-lbl">Инструкция</div>
            <div class="card-val">Загрузите рабочий Excel-файл</div>
            <div class="card-sub">
                Файл может содержать любое количество колонок (номенклатура, цена, количество…).<br>
                Обязательна колонка <b>Код ОКПД 2</b>.<br><br>
                В конец таблицы будут добавлены три колонки:<br>
                &nbsp;&nbsp;• <b>Наименование (ОКПД 2)</b> — из справочника ОКПД 2<br>
                &nbsp;&nbsp;• <b>В Приложении №1 (ПП 1875)</b> — ДА / НЕТ<br>
                &nbsp;&nbsp;• <b>В Приложении №2 (ПП 1875)</b> — ДА / НЕТ
            </div>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# TAB 2 — ПОИСК ПО КОДУ
# ══════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Поиск по коду ОКПД 2")
    st.caption("Введите полный или частичный код. Поддерживается иерархический поиск.")

    c_in, c_btn = st.columns([5, 1])
    with c_in:
        q_code = st.text_input(
            "Код", placeholder="Например: 13.2 или 17.23.13.191",
            label_visibility="collapsed", key="q_code"
        )
    with c_btn:
        go_code = st.button("Найти", type="primary", key="go_code")

    if go_code and q_code.strip():
        raw  = q_code.strip()
        code = clean_code(raw)

        if raw not in st.session_state.hist_code:
            st.session_state.hist_code.insert(0, raw)
            st.session_state.hist_code = st.session_state.hist_code[:10]

        fm, vm = hier_lookup(code, main_d)
        f1, _  = hier_lookup(code, app1_d)
        f2, _  = hier_lookup(code, app2_d)

        def badge(found: bool) -> str:
            c = 'b-yes' if found else 'b-no'
            t = '✅ ДА' if found else '❌ НЕТ'
            return f'<span class="badge {c}">{t}</span>'

        st.markdown(f"""
        <div class="card {'yes' if fm else 'no'}">
            <div class="card-lbl">📚 Справочник ОКПД 2 {badge(fm)}</div>
            <div class="card-val">{vm if fm else 'Код не найден в справочнике'}</div>
            <div class="card-sub">{code}</div>
        </div>
        <div class="card {'yes' if f1 else 'no'}">
            <div class="card-lbl">📋 Приложение №1 к ПП 1875 {badge(f1)}</div>
            <div class="card-val">{'Присутствует в Приложении №1' if f1 else 'Отсутствует в Приложении №1'}</div>
            <div class="card-sub">{code}</div>
        </div>
        <div class="card {'yes' if f2 else 'no'}">
            <div class="card-lbl">📋 Приложение №2 к ПП 1875 {badge(f2)}</div>
            <div class="card-val">{'Присутствует в Приложении №2' if f2 else 'Отсутствует в Приложении №2'}</div>
            <div class="card-sub">{code}</div>
        </div>
        """, unsafe_allow_html=True)

    if st.session_state.hist_code:
        st.markdown("---")
        st.markdown("#### 🕘 История")
        for i, h in enumerate(st.session_state.hist_code):
            ca, cb = st.columns([7, 1])
            with ca:
                st.markdown(f'<span class="mono">{h}</span>', unsafe_allow_html=True)
            with cb:
                if st.button("🗑️", key=f"dc{i}", type="secondary"):
                    st.session_state.hist_code.pop(i)
                    st.rerun()


# ══════════════════════════════════════════════════════════
# TAB 3 — ПОИСК ПО НАЗВАНИЮ
# ══════════════════════════════════════════════════════════
with tab3:
    st.markdown("### Поиск по наименованию в справочнике ОКПД 2")
    st.caption("Регистр не важен. Поиск подстроки. При нажатии Найти — автоматически проверяем Приложения.")

    c_in2, c_btn2 = st.columns([5, 1])
    with c_in2:
        q_name = st.text_input(
            "Наименование", placeholder="Например: аккумулятор или болт",
            label_visibility="collapsed", key="q_name"
        )
    with c_btn2:
        go_name = st.button("Найти", type="primary", key="go_name")

    if go_name and q_name.strip():
        qs = q_name.strip()
        q  = qs.lower()

        if qs not in st.session_state.hist_name:
            st.session_state.hist_name.insert(0, qs)
            st.session_state.hist_name = st.session_state.hist_name[:10]

        # Поиск в главном справочнике
        rows = [
            {'Код ОКПД 2': k, 'Наименование (ОКПД 2)': v}
            for k, v in main_d.items()
            if q in v.lower()
        ]

        if rows:
            st.success(f"Найдено в справочнике ОКПД 2: **{len(rows)}** записей")

            # Добавляем проверку по Приложениям
            enriched = []
            for r in rows:
                code = r['Код ОКПД 2']
                f1, _ = hier_lookup(code, app1_d)
                f2, _ = hier_lookup(code, app2_d)
                enriched.append({
                    'Код ОКПД 2':              code,
                    'Наименование (ОКПД 2)':   r['Наименование (ОКПД 2)'],
                    'В Приложении №1 (ПП 1875)': '✅ ДА' if f1 else '❌ НЕТ',
                    'В Приложении №2 (ПП 1875)': '✅ ДА' if f2 else '❌ НЕТ',
                })

            res_df = pd.DataFrame(enriched)
            st.dataframe(res_df, use_container_width=True, height=450)

            # Мини-статистика
            c1, c2, c3 = st.columns(3)
            c1.metric("Всего найдено", len(enriched))
            c2.metric("В Приложении №1", sum(1 for r in enriched if r['В Приложении №1 (ПП 1875)'] == '✅ ДА'))
            c3.metric("В Приложении №2", sum(1 for r in enriched if r['В Приложении №2 (ПП 1875)'] == '✅ ДА'))

            st.download_button(
                "⬇️ Скачать результат поиска (.xlsx)",
                data=to_excel(res_df),
                file_name=f"okpd2_search_{qs[:20]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error(f"❌ Совпадений по запросу «{qs}» не найдено в справочнике ОКПД 2.")

    if st.session_state.hist_name:
        st.markdown("---")
        st.markdown("#### 🕘 История")
        for i, h in enumerate(st.session_state.hist_name):
            ca, cb = st.columns([7, 1])
            with ca:
                st.markdown(f'<span class="mono">{h}</span>', unsafe_allow_html=True)
            with cb:
                if st.button("🗑️", key=f"dn{i}", type="secondary"):
                    st.session_state.hist_name.pop(i)
                    st.rerun()
